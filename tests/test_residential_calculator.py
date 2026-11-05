from __future__ import annotations

import unittest


REGIONAL_GRADES = {
    "natural_landscape": 2,
    "environmental_pollution": 2,
    "community_layout": 2,
    "building_density_appearance": 2,
    "outdoor_space_landscaping": 2,
    "repair_management": 2,
    "security_parking": 2,
    "sanitation_greening": 3,
    "main_road_distance": 3,
    "public_transport": 3,
    "school_district": 3,
    "medical_facilities": 3,
    "cultural_recreation": 3,
    "major_commercial": 3,
    "market_supermarket": 3,
}


def reference_project() -> dict:
    subject = {
        "address": "示例市滨江区示范路88号示例花园5幢1单元1801室",
        "report_title": "示例花园5幢1单元1801室住宅房地产市场价值评估",
        "report_number": "示例住宅估字[2026]001号",
        "dispatch_number": "TEST-2026-001",
        "valuation_date": "2026-08-28",
        "owner": "测试产权人",
        "certificate_no": "苏（2026）示例不动产权第TEST001号",
        "co_ownership": "单独所有",
        "real_estate_unit_no": "TEST-UNIT-001",
        "land_area": 32543.20,
        "land_use": "城镇住宅用地",
        "rights_type": "国有建设用地使用权/房屋所有权",
        "rights_nature": "出让/安置房",
        "rights_term": "2018年04月16日起2088年04月15日止",
        "parcel_no": "TEST-PARCEL-001",
        "location_boundaries": "东临测试路，西临示例路，南临规划路，北临公园路。",
        "location_description": "估价对象位于测试居住片区，周边交通和生活配套完善。",
        "location_analysis": "测试区域居住环境较好，区位条件对估价对象价值产生积极影响。",
        "area": 121.43,
        "structure": "钢混结构",
        "certificate_floor_display": "18/19",
        "below_ground_floors": 1,
        "above_ground_total_floors": 18,
        "floor": 18,
        "build_year": 2019,
        "orientation": "二室朝南",
        "east_west": "东边户",
        "building_type": "high_rise",
        "property_nature": "安置房",
        "area_factor_pct": 100,
        "decoration_unit_price": 0,
        "decoration_newness": 0,
        "appurtenance_description": "/",
        "appurtenance_unit_price": 0,
    }
    comparables = [
        {
            "name": "可比实例A",
            "source": "贝壳房产网",
            "address": "示例市滨江区示范路88号示例花园1幢2单元1606室",
            "area": 62.85,
            "structure": "钢混结构",
            "above_ground_total_floors": 18,
            "floor": 16,
            "build_year": 2019,
            "orientation": "一室朝南",
            "east_west": "中间户",
            "building_type": "high_rise",
            "property_nature": "安置房",
            "area_factor_pct": 102,
            "decoration_unit_price": 500,
            "decoration_newness": 0.8,
            "appurtenance_description": "/",
            "appurtenance_unit_price": 0,
            "transaction_date": "2026-08-01",
            "transaction_unit_price": 3341,
            "transaction_condition_factor": 1,
            "market_condition_factor": 1,
        },
        {
            "name": "可比实例B",
            "source": "E路平安信息系统",
            "address": "示例市滨江区示范路88号示例花园4幢1单元1601室",
            "area": 102.04,
            "structure": "钢混结构",
            "above_ground_total_floors": 18,
            "floor": 16,
            "build_year": 2019,
            "orientation": "二室朝南",
            "east_west": "东边户",
            "building_type": "high_rise",
            "property_nature": "安置房",
            "area_factor_pct": 100,
            "decoration_unit_price": 500,
            "decoration_newness": 0.8,
            "appurtenance_description": "/",
            "appurtenance_unit_price": 0,
            "transaction_date": "2026-06-02",
            "transaction_unit_price": 3214,
            "transaction_condition_factor": 1,
            "market_condition_factor": 1,
        },
        {
            "name": "可比实例C",
            "source": "E路平安信息系统",
            "address": "示例市滨江区示范路88号示例花园3幢1单元903室",
            "area": 121.22,
            "structure": "钢混结构",
            "above_ground_total_floors": 18,
            "floor": 9,
            "build_year": 2019,
            "orientation": "二室朝南",
            "east_west": "中间户",
            "building_type": "high_rise",
            "property_nature": "安置房",
            "area_factor_pct": 100,
            "decoration_unit_price": 500,
            "decoration_newness": 0.8,
            "appurtenance_description": "/",
            "appurtenance_unit_price": 0,
            "transaction_date": "2026-05-18",
            "transaction_unit_price": 3217,
            "transaction_condition_factor": 1,
            "market_condition_factor": 1,
        },
    ]
    return {
        "subject": subject,
        "comparables": comparables,
        "regional_grades": {
            "subject": dict(REGIONAL_GRADES),
            "comparables": [dict(REGIONAL_GRADES) for _ in comparables],
        },
        "adjustment_notes": {},
    }


class ResidentialCalculatorGoldenTests(unittest.TestCase):
    def test_reference_project_matches_2779_and_337454(self):
        from residential_calculator import calculate_residential_comparison

        result = calculate_residential_comparison(reference_project())

        self.assertEqual(result["estimated_unit_price"], "2779")
        self.assertEqual(result["estimated_total"], "337454")
        self.assertEqual(
            result["total_uppercase"],
            "人民币叁拾叁万柒仟肆佰伍拾肆元整",
        )
        self.assertEqual(
            [item["comparable_unit_price"] for item in result["comparables"]],
            ["2824", "2734", "2778"],
        )

    def test_certificate_and_correction_floors_are_separate(self):
        from residential_calculator import calculate_residential_comparison

        result = calculate_residential_comparison(reference_project())

        subject = result["subject"]
        self.assertEqual(subject["certificate_floor_display"], "18/19")
        self.assertEqual(subject["above_ground_total_floors"], 18)
        self.assertEqual(subject["floor_correction_display"], "18/18")
        self.assertEqual(subject["level_factor_pct"], "99.40")
        self.assertEqual(subject["factor_breakdown"]["level"], "99.40")
        self.assertEqual(result["comparables"][0]["factor_breakdown"]["level"], "102.00")

    def test_non_neutral_p1_and_p2_affect_result(self):
        from residential_calculator import calculate_residential_comparison

        project = reference_project()
        project["comparables"][0]["transaction_condition_factor"] = 0.98
        project["comparables"][0]["market_condition_factor"] = 1.03

        result = calculate_residential_comparison(project)

        self.assertEqual(result["comparables"][0]["p1"], "0.980")
        self.assertEqual(result["comparables"][0]["p2"], "1.030")
        self.assertNotEqual(result["estimated_unit_price"], "2779")

    def test_non_neutral_factor_without_note_adds_warning(self):
        from residential_calculator import calculate_residential_comparison

        project = reference_project()
        project["comparables"][0]["transaction_condition_factor"] = 0.98
        result = calculate_residential_comparison(project)

        self.assertTrue(
            any("可比实例A交易情况修正未填写依据" in warning for warning in result["warnings"])
        )


class ResidentialCalculatorValidationTests(unittest.TestCase):
    def test_applies_multi_story_and_elevator_floor_tables(self):
        from residential_calculator import calculate_residential_comparison

        project = reference_project()
        project["subject"].update(
            {"building_type": "multi_story", "floor": 6, "above_ground_total_floors": 6}
        )
        for item, floor in zip(project["comparables"], (5, 4, 3)):
            item.update(
                {"building_type": "multi_story", "floor": floor, "above_ground_total_floors": 6}
            )
        result = calculate_residential_comparison(project)
        self.assertEqual(
            [result["subject"]["factor_breakdown"]["level"]]
            + [item["factor_breakdown"]["level"] for item in result["comparables"]],
            ["93.00", "98.00", "103.00", "105.00"],
        )

        project = reference_project()
        project["subject"].update(
            {
                "building_type": "elevator_multi_story",
                "floor": 6,
                "above_ground_total_floors": 6,
            }
        )
        for item, floor in zip(project["comparables"], (5, 4, 3)):
            item.update(
                {
                    "building_type": "elevator_multi_story",
                    "floor": floor,
                    "above_ground_total_floors": 6,
                }
            )
        result = calculate_residential_comparison(project)
        self.assertEqual(
            [result["subject"]["factor_breakdown"]["level"]]
            + [item["factor_breakdown"]["level"] for item in result["comparables"]],
            ["106.00", "106.00", "105.00", "104.00"],
        )

    def test_rejects_missing_case_regional_grades(self):
        from residential_calculator import (
            ResidentialCalculationError,
            calculate_residential_comparison,
        )

        project = reference_project()
        project["regional_grades"]["comparables"].pop()

        with self.assertRaisesRegex(ResidentialCalculationError, "可比实例C"):
            calculate_residential_comparison(project)

    def test_rejects_non_positive_transaction_unit_price(self):
        from residential_calculator import (
            ResidentialCalculationError,
            calculate_residential_comparison,
        )

        project = reference_project()
        project["comparables"][0]["transaction_unit_price"] = 0

        with self.assertRaisesRegex(ResidentialCalculationError, "可比实例A成交单价"):
            calculate_residential_comparison(project)


if __name__ == "__main__":
    unittest.main()
