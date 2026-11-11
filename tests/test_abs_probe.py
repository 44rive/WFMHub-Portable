"""Synthetic traversal/privacy tests plus a real Windows COM smoke test."""
import ctypes
import json
import os
from pathlib import Path
import tempfile
import threading
import unittest

from miniapps.bo_live.abs_probe import GUID, VARIANT, WindowsUIAutomation, describe, inspect_tree, main


class FakeTree:
    def __init__(self):
        self.released = []

    def read(self, element):
        return {"name": describe({1:"Task priority",2:"PRIVATE-CASE-000123",3:"Scheduled for"}[element]),
                "grid": element == 1, "grid_item": element == 2}

    def adjacent(self, element, child=False):
        return ({1:2} if child else {2:3}).get(element)

    def release(self, element):
        self.released.append(element)


class ProbeTests(unittest.TestCase):
    def test_guid_and_win64_variant_layout(self):
        self.assertEqual(ctypes.sizeof(GUID), 16)
        self.assertEqual(ctypes.sizeof(VARIANT), 24)
        self.assertEqual(GUID.parse('FF48DBA4-60EF-4201-AA87-54103EEF594E').a, 0xFF48DBA4)

    def test_arbitrary_text_is_redacted_and_truncation_preserved(self):
        value=describe('PRIVATE CUSTOMER / CASE 00012...')
        self.assertTrue(value['redacted'])
        self.assertTrue(value['ends_in_ellipsis'])
        self.assertNotIn('PRIVATE', json.dumps(value))
        self.assertEqual(describe(' Scheduled for ')['header'], 'Scheduled for')

    def test_tree_preserves_positions_but_never_claims_complete_backlog(self):
        tree=FakeTree(); report=inspect_tree(tree, 1)
        self.assertEqual(report['summary']['nodes_scanned'],3)
        self.assertEqual(report['summary']['positioned_cells'],1)
        self.assertEqual(report['summary']['headers_found'],['Scheduled for','Task priority'])
        self.assertEqual(report['summary']['full_backlog_coverage'],'not_verified')
        self.assertEqual(sorted(tree.released),[1,2,3])
        self.assertNotIn('PRIVATE',json.dumps(report))

    def test_node_limit_releases_unvisited_references(self):
        tree=FakeTree(); report=inspect_tree(tree,1,max_nodes=1)
        self.assertTrue(report['summary']['limit_reached'])
        self.assertEqual(report['summary']['nodes_scanned'],1)
        self.assertEqual(tree.released,[1,2])

    def test_password_nodes_are_not_traversed(self):
        tree=FakeTree(); tree.read=lambda element:{'password_control':True}
        report=inspect_tree(tree,1)
        self.assertEqual(report['summary']['nodes_scanned'],1)
        self.assertEqual(tree.released,[1])

    @unittest.skipIf(os.name=='nt','Non-Windows failure path')
    def test_non_windows_reports_unavailability_instead_of_fake_capture(self):
        with tempfile.TemporaryDirectory() as folder:
            self.assertEqual(main(['--output',folder]),1)
            report=json.loads(next(Path(folder).glob('*.json')).read_text())
            self.assertEqual(report['status'],'probe_failed')
            self.assertTrue(report['values_redacted'])
            self.assertNotIn('rows',report)


@unittest.skipUnless(os.name=='nt','Requires native Windows COM')
class NativeProbeTests(unittest.TestCase):
    def test_native_com_reads_a_synthetic_window_and_releases_bstr(self):
        from ctypes import wintypes as W
        C=ctypes
        user=C.WinDLL('user32',use_last_error=True)
        kernel=C.WinDLL('kernel32',use_last_error=True)
        user.CreateWindowExW.argtypes=[W.DWORD,W.LPCWSTR,W.LPCWSTR,W.DWORD,C.c_int,C.c_int,C.c_int,C.c_int,W.HWND,W.HMENU,W.HINSTANCE,C.c_void_p]
        user.CreateWindowExW.restype=W.HWND
        user.DestroyWindow.argtypes=[W.HWND]; user.DestroyWindow.restype=W.BOOL
        user.GetMessageW.argtypes=[C.POINTER(W.MSG),W.HWND,W.UINT,W.UINT]; user.GetMessageW.restype=W.BOOL
        user.DispatchMessageW.argtypes=[C.POINTER(W.MSG)]; user.DispatchMessageW.restype=C.c_ssize_t
        user.PostThreadMessageW.argtypes=[W.DWORD,W.UINT,W.WPARAM,W.LPARAM]; user.PostThreadMessageW.restype=W.BOOL
        kernel.GetCurrentThreadId.restype=W.DWORD
        ready=threading.Event(); shared={}
        def owner():
            shared['thread']=kernel.GetCurrentThreadId()
            hwnd=user.CreateWindowExW(0,'STATIC','Task priority',0x00CF0000,0,0,200,100,None,None,None,None)
            shared['hwnd']=hwnd
            message=W.MSG()
            # Create a message queue before the test can post its shutdown signal.
            user.PeekMessageW.argtypes=[C.POINTER(W.MSG),W.HWND,W.UINT,W.UINT,W.UINT]
            user.PeekMessageW(C.byref(message),None,0,0,0)
            ready.set()
            if hwnd:
                while user.GetMessageW(C.byref(message),None,0,0)>0:
                    user.DispatchMessageW(C.byref(message))
                user.DestroyWindow(hwnd)
        thread=threading.Thread(target=owner,daemon=True); thread.start()
        self.assertTrue(ready.wait(5)); self.assertTrue(shared.get('hwnd'))
        automation=None; element=None
        try:
            automation=WindowsUIAutomation()
            element=automation.element(shared['hwnd'])
            for unused in range(3):
                self.assertEqual(automation.property(element,30005),'Task priority')
            self.assertIsInstance(automation.property(element,30003),int)
        finally:
            if automation:
                automation.release(element); automation.close()
            user.PostThreadMessageW(shared['thread'],0x0012,0,0)
            thread.join(timeout=5)
        self.assertFalse(thread.is_alive())


if __name__=='__main__':
    unittest.main()
