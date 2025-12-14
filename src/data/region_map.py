from __future__ import annotations

from itertools import chain

REGION_DEFINITIONS = {
    "North America": """
        CAN USA GRL BMU
    """,
    "Latin America & Caribbean": """
        ABW AIA ATG ARG BHS BRB BLZ BOL BRA CHL COL CRI CUB CUW CYM
        DMA DOM ECU GLP GRD GTM GUF GUY HND HTI JAM KNA LCA MEX MSR MTQ
        NIC PAN PER PRI PRY SUR TTO URY VEN VCT VGB
    """,
    "Europe & Central Asia": """
        ALA ALB AND ARM AUT AZE BEL BGR BIH BLR CHE CYP CZE DEU DNK ESP EST
        FIN FRA GBR GEO GRC HRV HUN IRL ISL ITA KAZ KGZ LTU LUX LVA MDA MKD
        MLT MNE NLD NOR POL PRT ROU RUS SMR SRB SVK SVN SWE TUR UKR UZB
    """,
    "Middle East & North Africa": """
        ARE BHR DZA EGY IRN IRQ ISR JOR KWT LBN LBY MAR MRT OMN PSE QAT
        SAU SDN SYR TUN YEM
    """,
    "Sub-Saharan Africa": """
        AGO BDI BEN BFA BWA CIV CMR COD COG COM CPV DJI ETH GAB GHA GIN
        GMB GNB GNQ KEN LBR LSO MDG MLI MOZ MUS MWI NAM NER NGA RWA SEN
        SLE SOM SSD STP SWZ TCD TGO TZA UGA ZAF ZMB ZWE REU
    """,
    "South Asia": """
        AFG BGD BTN IND LKA MDV NPL PAK
    """,
    "East Asia & Pacific": """
        AUS BRN CHN COK FJI FSM GUM HKG IDN JPN KHM KIR KOR LAO MAC MMR
        MNG MYS NCL NZL PHL PNG PLW SGP SLB THA TLS TON VNM VUT WSM
    """,
}

AGGREGATE_OVERRIDES = {
    "NAC": "North America",
    "SAS": "South Asia",
    "WLD": "Global",
}

ISO_REGION_MAP = {}
for region, blob in REGION_DEFINITIONS.items():
    codes = [code.strip() for code in blob.split() if code.strip()]
    ISO_REGION_MAP.update({code: region for code in codes})

# Ensure aggregate placeholders map cleanly
ISO_REGION_MAP.update(AGGREGATE_OVERRIDES)

# Additional overrides for codes not included above
ISO_REGION_MAP.setdefault("PYF", "East Asia & Pacific")
ISO_REGION_MAP.setdefault("REU", "Sub-Saharan Africa")
ISO_REGION_MAP.setdefault("SLV", "Latin America & Caribbean")
ISO_REGION_MAP.setdefault("SYC", "Sub-Saharan Africa")
ISO_REGION_MAP.setdefault("TJK", "Europe & Central Asia")

REGION_DISPLAY_ORDER = [
    "Global",
    "North America",
    "Latin America & Caribbean",
    "Europe & Central Asia",
    "Middle East & North Africa",
    "Sub-Saharan Africa",
    "South Asia",
    "East Asia & Pacific",
]

REGION_FALLBACK = "Other / Unmapped"
