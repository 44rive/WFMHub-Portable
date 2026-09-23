import sys

try:
    from .cli import main
except (ImportError, ModuleNotFoundError, OSError) as exc:
    print("\nWFMHUB COULD NOT START")
    print(f"Python : {sys.executable}")
    print(f"Cause  : {type(exc).__name__}: {exc}")
    print("\nThe portable runtime is incomplete or Windows blocked one of its files.")
    print("Download the Windows ZIP from GitHub Releases and use Extract All into a new local folder.")
    print("Do not merge it over an older WFMHub folder and do not use GitHub's Source code ZIP.")
    raise SystemExit(78) from None


if __name__ == "__main__":
    raise SystemExit(main())
