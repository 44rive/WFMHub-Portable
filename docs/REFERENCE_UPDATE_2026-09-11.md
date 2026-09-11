# Operations reference update — 2026-09-11

Authority inspected: `attachments/REF.xlsx` plus the four September Verint
forecast exports and the `Activities`/`StartEndTimes` extracts. Source files
remain untouched and are not shipped in the portable release.

## RSA Netherlands Flash difference

The previous shipped RSA NL Flash allowlist had 30 queues. The supplied
`Queues Ref` sheet marks 23 exact queues `SL Related = Y`; those 23 are now the
Flash total. Eight `N` queues remain eligible for mapped clean data but do not
enter RSA NL SL. Two `Y` queues were absent from the old map and were added.

Added to mapping and Flash:

- `APBN_AMS_MOBILITY_INSURAN_AllianzNetherlandsAlarm_NL`
- `APBN_AMS_RSA_OEM_Toyota_NL`

Excluded from the RSA NL Flash because `SL Related = N`:

- `APBN_AMS_MOBILITY_CRM_ConciergeBBH_NL`
- `APBN_AMS_MOBILITY_LEASE-RE_FollowUp_NL`
- `APBN_AMS_MOBILITY_PROVIDER_Interco_NL`
- `APBN_AMS_MOBILITY_PROVIDER_Local_NL`
- `APBN_AMS_MOBILITY_PROVIDER_Morocco_NL`
- `APBN_AMS_RSA_CRM_ConciergeBBH_NL`
- `APBN_AMS_RSA_Ford_Dealers_NL`
- `APBN_AMS_RSA_PROVIDER_All_NL`

`APBN_AMS_MOBILITY_NIGHT_NightShift_NL` is absent from the new reference and is
therefore also excluded from the Flash allowlist. It is not deleted from stored
history merely because it is outside this presentation scope.

The exact 23-queue RSA NL set is:

1. `APBN_AMS_MOBILITY_BIKE_Bike_NL`
2. `APBN_AMS_MOBILITY_INSURAN_AllianzNetherlandsAlarm_NL`
3. `APBN_AMS_MOBILITY_INSURAN_AllianzNetherlandsOEM_NL`
4. `APBN_AMS_MOBILITY_INSURAN_FollowUp_NL`
5. `APBN_AMS_MOBILITY_INSURAN_Front_EN`
6. `APBN_AMS_MOBILITY_INSURAN_Front_NL`
7. `APBN_AMS_MOBILITY_INTERNAT_FollowUp_NL`
8. `APBN_AMS_MOBILITY_INTERNAT_Front_NL`
9. `APBN_AMS_MOBILITY_LEASE-RE_Front_NL`
10. `APBN_AMS_MOBILITY_OEM-CONS_FollowUp_NL`
11. `APBN_AMS_MOBILITY_OEM-CONS_Front_NL`
12. `APBN_AMS_MOBILITY_VIP_FollowUp_NL`
13. `APBN_AMS_MOBILITY_VIP_Front_NL`
14. `APBN_AMS_RSA_Ford_Assistance_NL`
15. `APBN_AMS_RSA_INSURAN_All_EN`
16. `APBN_AMS_RSA_INSURAN_All_NL`
17. `APBN_AMS_RSA_INSURAN_AzN_NL`
18. `APBN_AMS_RSA_INTERNAT_All_EN`
19. `APBN_AMS_RSA_INTERNAT_All_NL`
20. `APBN_AMS_RSA_OEM_All_EN`
21. `APBN_AMS_RSA_OEM_All_NL`
22. `APBN_AMS_RSA_VIP_All_NL`
23. `APBN_AMS_RSA_OEM_Toyota_NL`

The new reference supplies no replacement exact queue set for RSA BE, Ford NL
or OEM, so those reviewed profile memberships were not guessed or broadened.

## Forecast contract

The four supplied September forecast files are native 15-minute exports. Their
schema contains queue, date, time, interval and volume only. WFMHub therefore:

- preserves each 15-minute volume row in `mart.forecast_interval`;
- sums the four quarters for hourly volume views;
- leaves FTE Required, staffing forecast, forecast SL and forecast AHT blank
  because those fields are absent from the supplied files.

No staffing requirement is inferred from volume.

## Attendance and final absence contract

- StartEndTimes remains the preferred shift boundary.
- Agent Status remains the actual presence, RTM, attendance, break and gap
  source. The exact AUX and Status labels in `REF.xlsx` already match the
  shipped `default_rules.toml` reference rows.
- LILO remains fallback/control when Agent Status coverage is insufficient.
- Activities is used only for final post-day absence/shrinkage classifications
  and as an explicit schedule-boundary fallback. It never creates presence.
- Empty shifts, unmapped Activities, PTO/Away missing from Activities and
  partially corrected observed gaps remain visible review states.
