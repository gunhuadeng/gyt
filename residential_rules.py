from __future__ import annotations


REGIONAL_FACTORS = (
    {"key": "natural_landscape", "label": "自然景观", "scores": (10, 9, 8, 7, 6)},
    {"key": "environmental_pollution", "label": "环境污染", "scores": (8, 7.5, 7, 6.5, 6)},
    {"key": "community_layout", "label": "建筑或小区布局", "scores": (6, 5.5, 5, 4.5, 4)},
    {"key": "building_density_appearance", "label": "建筑密度、外型等", "scores": (5, 4.5, 4, 3.5, 3)},
    {"key": "outdoor_space_landscaping", "label": "室外公共活动空间与绿化景观", "scores": (7, 6.5, 6, 5.5, 5)},
    {"key": "repair_management", "label": "房屋、设施维修与管理", "scores": (6, 5.5, 5, 4.5, 4)},
    {"key": "security_parking", "label": "小区治安、道路停车管理", "scores": (6, 5.5, 5, 4.5, 4)},
    {"key": "sanitation_greening", "label": "环卫与绿化管理", "scores": (6, 5.5, 5, 4.5, 4)},
    {"key": "main_road_distance", "label": "离市区主干道的距离", "scores": (12, 11, 10, 9, 8)},
    {"key": "public_transport", "label": "公交线路情况", "scores": (12, 11, 10, 9, 8)},
    {"key": "school_district", "label": "所在学区学校情况", "scores": (6, 5.5, 5, 4.5, 4)},
    {"key": "medical_facilities", "label": "医院及医疗机构分布情况", "scores": (6, 5.5, 5, 4.5, 4)},
    {"key": "cultural_recreation", "label": "文体娱乐设施", "scores": (6, 5.5, 5, 4.5, 4)},
    {"key": "major_commercial", "label": "重要商业配套设施", "scores": (10, 9, 8, 7, 6)},
    {"key": "market_supermarket", "label": "菜场超市", "scores": (14, 13, 12, 11, 10)},
)

STRUCTURE_FACTORS = {
    "钢混结构": 108,
    "砖混结构": 100,
    "混合结构": 100,
    "砖木结构": 95,
}

ORIENTATION_FACTORS = {
    "北": 96,
    "西": 98,
    "东": 99,
    "一室朝南": 100,
    "二室朝南": 101,
    "三室朝南": 102,
}

EAST_WEST_FACTORS = {
    "东边户": 106,
    "中间户": 100,
    "西边户": 104,
}

MULTI_STORY_LEVELS = {
    (1, 1): 100, (1, 2): 100, (1, 3): 100, (1, 4): 100, (1, 5): 100, (1, 6): 100, (1, 7): 100,
    (2, 2): 100, (2, 3): 101, (2, 4): 102, (2, 5): 102, (2, 6): 102, (2, 7): 102,
    (3, 3): 100, (3, 4): 103, (3, 5): 104, (3, 6): 105, (3, 7): 105,
    (4, 4): 95, (4, 5): 101, (4, 6): 103, (4, 7): 105,
    (5, 5): 94, (5, 6): 98, (5, 7): 100,
    (6, 6): 93, (6, 7): 98,
    (7, 7): 91,
}

ELEVATOR_MULTI_STORY_LEVELS = {
    (1, 4): 100, (1, 5): 100, (1, 6): 100, (1, 7): 100,
    (2, 4): 102, (2, 5): 102, (2, 6): 102, (2, 7): 102,
    (3, 4): 104, (3, 5): 104, (3, 6): 104, (3, 7): 104,
    (4, 4): 104, (4, 5): 105, (4, 6): 105, (4, 7): 105,
    (5, 5): 105, (5, 6): 106, (5, 7): 106,
    (6, 6): 106, (6, 7): 107,
    (7, 7): 107,
}


def high_rise_level_factor(floor: int, total_floors: int) -> float:
    if total_floors < 6 or floor < 1 or floor > total_floors:
        raise ValueError("高层楼层数据超出规则范围")
    if floor <= 6:
        return (97.5, 98.0, 98.5, 99.0, 99.5, 100.0)[floor - 1]
    if floor == 7:
        return 100.2
    if floor == total_floors:
        return round(100 + (total_floors - 6) * 0.2 - 3, 2)
    return round(100 + (floor - 6) * 0.2, 2)
