# Governed queue catalog

Last reconciled: `2026-09-14`

This document explains the queue boundary; the executable authority remains
`config/service_profiles.toml` plus `config/queue_mapping.csv`. The Manager
Workbench **Govern > Mappings & Rules > Queue register** displays every current
mapped queue with its primary data scope, exact Flash views, service status,
workforce owner and overlap rule.

## Service and capacity are different decisions

- **Service-related queue** means the exact queue is in at least one active
  `flash_queues` allowlist. It contributes Call-by-Call counters to that Flash.
- **Non-service queue** remains mapped and available in clean data and queue
  diagnostics, but it does not contribute to a Flash total.
- **Workforce owner** is a capacity identity. It comes from the primary service
  scope and the profile's explicit staffing link; it is not inferred from the
  queue name or from every Flash in which the queue appears.
- A dual-view queue is counted in both exact configured service views. That
  overlap does not duplicate or move its scheduled workforce.

Current exact Flash membership counts are RSA NL `23`, RSA BE `43`, Ford NL
`6`, and OEM `3`. Membership counts include intentional cross-view overlap.

## Ford VL / DE ownership

The supplied service references place these four queues in both the combined
RSA BE Flash and Ford NL Flash. Their primary data scope remains `Ford NL` and
their staffing workforce remains `Ford Dutch`:

| Queue | Flash views | Primary data scope | Workforce owner |
|---|---|---|---|
| APBN_BRU_MOBILITY_Ford_Assistance_DE | RSA BE + FORD NL | Ford NL | Ford Dutch |
| APBN_BRU_MOBILITY_Ford_Assistance_VL | RSA BE + FORD NL | Ford NL | Ford Dutch |
| APBN_BRU_MOBILITY_Ford_Dealers_VL | RSA BE + FORD NL | Ford NL | Ford Dutch |
| APBN_LUX_MOBILITY_Ford_Assistance_DE | RSA BE + FORD NL | Ford NL | Ford Dutch |

This is deliberate: RSA BE service is a combined service result, while future
capacity is prepared at separate workforce planning grains.

## RSA NL REF.xlsx reconciliation

All 31 rows in `attachments/REF.xlsx` are mapped. The 23 `Y` rows are in the
exact RSA NL Flash allowlist and none of the eight `N` rows are admitted.

| Queue | SL related |
|---|---:|
| APBN_AMS_MOBILITY_BIKE_Bike_NL | Y |
| APBN_AMS_MOBILITY_CRM_ConciergeBBH_NL | N |
| APBN_AMS_MOBILITY_INSURAN_AllianzNetherlandsAlarm_NL | Y |
| APBN_AMS_MOBILITY_INSURAN_AllianzNetherlandsOEM_NL | Y |
| APBN_AMS_MOBILITY_INSURAN_FollowUp_NL | Y |
| APBN_AMS_MOBILITY_INSURAN_Front_EN | Y |
| APBN_AMS_MOBILITY_INSURAN_Front_NL | Y |
| APBN_AMS_MOBILITY_INTERNAT_FollowUp_NL | Y |
| APBN_AMS_MOBILITY_INTERNAT_Front_NL | Y |
| APBN_AMS_MOBILITY_LEASE-RE_FollowUp_NL | N |
| APBN_AMS_MOBILITY_LEASE-RE_Front_NL | Y |
| APBN_AMS_MOBILITY_OEM-CONS_FollowUp_NL | Y |
| APBN_AMS_MOBILITY_OEM-CONS_Front_NL | Y |
| APBN_AMS_MOBILITY_PROVIDER_Interco_NL | N |
| APBN_AMS_MOBILITY_PROVIDER_Local_NL | N |
| APBN_AMS_MOBILITY_PROVIDER_Morocco_NL | N |
| APBN_AMS_MOBILITY_VIP_FollowUp_NL | Y |
| APBN_AMS_MOBILITY_VIP_Front_NL | Y |
| APBN_AMS_RSA_CRM_ConciergeBBH_NL | N |
| APBN_AMS_RSA_Ford_Assistance_NL | Y |
| APBN_AMS_RSA_Ford_Dealers_NL | N |
| APBN_AMS_RSA_INSURAN_All_EN | Y |
| APBN_AMS_RSA_INSURAN_All_NL | Y |
| APBN_AMS_RSA_INSURAN_AzN_NL | Y |
| APBN_AMS_RSA_INTERNAT_All_EN | Y |
| APBN_AMS_RSA_INTERNAT_All_NL | Y |
| APBN_AMS_RSA_OEM_All_EN | Y |
| APBN_AMS_RSA_OEM_All_NL | Y |
| APBN_AMS_RSA_PROVIDER_All_NL | N |
| APBN_AMS_RSA_VIP_All_NL | Y |
| APBN_AMS_RSA_OEM_Toyota_NL | Y |

Never add or move a queue by substring, suffix, language marker or an assumed
agent team. Update the reviewed references and both governed configuration files
with a regression test when the business scope changes.
