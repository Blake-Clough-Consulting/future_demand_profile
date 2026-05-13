import unittest

import pandas as pd

from csv_processor import (
    DATE_OF_AGGREGATION_COLUMN,
    DERIVED_PROFILE_COLUMN,
    FLOW_RUN_DATE_COLUMN,
    IMPORT_EXPORT_INDICATOR_COLUMN,
    RAW_PROFILE_COLUMN,
    SETTLEMENT_RUN_TYPE_COLUMN,
    SIGNED_PROFILE_COLUMN,
    add_signed_meter_volume,
    scale_profile_to_net_demand_setpoints,
    select_preferred_settlement_runs,
)


class SignedMeterVolumeTests(unittest.TestCase):
    def test_import_and_export_sign_conventions_are_applied(self):
        df = pd.DataFrame(
            {
                RAW_PROFILE_COLUMN: [10.0, 20.0, -30.0],
                IMPORT_EXPORT_INDICATOR_COLUMN: ["I", "E", "E"],
            }
        )

        signed_df, summary = add_signed_meter_volume(df, verbose=False)

        self.assertEqual(signed_df[SIGNED_PROFILE_COLUMN].tolist(), [10.0, -20.0, -30.0])
        self.assertEqual(summary["import_rows"], 1)
        self.assertEqual(summary["export_rows"], 2)
        self.assertEqual(summary["positive_export_rows_corrected"], 1)
        self.assertEqual(summary["negative_export_rows_preserved"], 1)

    def test_unknown_and_missing_indicators_are_left_unchanged(self):
        df = pd.DataFrame(
            {
                RAW_PROFILE_COLUMN: [40.0, 50.0, 60.0],
                IMPORT_EXPORT_INDICATOR_COLUMN: ["X", None, ""],
            }
        )

        signed_df, summary = add_signed_meter_volume(df, verbose=False)

        self.assertEqual(signed_df[SIGNED_PROFILE_COLUMN].tolist(), [40.0, 50.0, 60.0])
        self.assertEqual(summary["unknown_indicator_rows"], 1)
        self.assertEqual(summary["missing_indicator_rows"], 2)

    def test_scaling_uses_signed_meter_volume_not_raw_meter_volume(self):
        df = pd.DataFrame(
            {
                RAW_PROFILE_COLUMN: [200.0, 100.0],
                SIGNED_PROFILE_COLUMN: [-10.0, 10.0],
            }
        )
        net_demand_setpoints = {
            "winter_peak": {"net_demand": 20.0},
            "summer_min_am": {"net_demand": 0.0, "label": "Summer Min AM"},
            "summer_min_pm": {"net_demand": 5.0, "label": "Summer Min PM"},
        }

        derived_df, scaling_summary = scale_profile_to_net_demand_setpoints(df, net_demand_setpoints)

        self.assertEqual(derived_df[DERIVED_PROFILE_COLUMN].tolist(), [0.0, 20.0])
        self.assertEqual(scaling_summary["source_min"], -10.0)
        self.assertEqual(scaling_summary["source_max"], 10.0)


class SettlementRunPriorityTests(unittest.TestCase):
    def test_prefers_reconciliation_runs_in_priority_order(self):
        df = pd.DataFrame(
            {
                "GSP Id": ["NAIR_P", "NAIR_P", "NAIR_P", "NAIR_P"],
                SETTLEMENT_RUN_TYPE_COLUMN: ["SF", "R1", "R3", "R2"],
                FLOW_RUN_DATE_COLUMN: [20230101, 20230201, 20230401, 20230301],
                DATE_OF_AGGREGATION_COLUMN: [20230101, 20230201, 20230401, 20230301],
                RAW_PROFILE_COLUMN: [1.0, 2.0, 4.0, 3.0],
            },
            index=pd.to_datetime(["2022-01-01 00:00"] * 4),
        )

        selected_df, summary = select_preferred_settlement_runs(df, verbose=False)

        self.assertEqual(len(selected_df), 1)
        self.assertEqual(selected_df.iloc[0][SETTLEMENT_RUN_TYPE_COLUMN], "R3")
        self.assertEqual(selected_df.iloc[0][RAW_PROFILE_COLUMN], 4.0)
        self.assertEqual(summary["selected_run_type_counts"], {"R3": 1})
        self.assertEqual(summary["duplicate_rows_removed"], 3)

    def test_falls_back_to_best_available_run(self):
        df = pd.DataFrame(
            {
                "GSP Id": ["NAIR_P", "NAIR_P"],
                SETTLEMENT_RUN_TYPE_COLUMN: ["SF", "R1"],
                RAW_PROFILE_COLUMN: [1.0, 2.0],
            },
            index=pd.to_datetime(["2022-01-01 00:00", "2022-01-01 00:00"]),
        )

        selected_df, summary = select_preferred_settlement_runs(df, verbose=False)

        self.assertEqual(selected_df.iloc[0][SETTLEMENT_RUN_TYPE_COLUMN], "R1")
        self.assertEqual(summary["selected_run_type_counts"], {"R1": 1})
        self.assertEqual(summary["fallback_rows"], 1)

    def test_keeps_same_datetime_for_different_gsps(self):
        df = pd.DataFrame(
            {
                "GSP Id": ["A", "A", "B", "B"],
                SETTLEMENT_RUN_TYPE_COLUMN: ["SF", "R2", "R1", "R3"],
                RAW_PROFILE_COLUMN: [1.0, 2.0, 3.0, 4.0],
            },
            index=pd.to_datetime(["2022-01-01 00:00"] * 4),
        )

        selected_df, summary = select_preferred_settlement_runs(df, verbose=False)

        self.assertEqual(len(selected_df), 2)
        self.assertEqual(selected_df.sort_values("GSP Id")[SETTLEMENT_RUN_TYPE_COLUMN].tolist(), ["R2", "R3"])
        self.assertEqual(summary["selected_run_type_counts"], {"R2": 1, "R3": 1})


if __name__ == "__main__":
    unittest.main()
