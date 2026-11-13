"""
竞价行为分析模块
对每台机组计算指标（基于最新年度），使用多年度数据进行历史分析。
"""
import csv
import math
import os
from typing import Dict, List, Optional, Any, Tuple
from dataclasses import dataclass, field

# 尝试导入 pandas 以支持 Excel 读取
try:
    import pandas as pd
    PANDAS_AVAILABLE = True
except ImportError:
    PANDAS_AVAILABLE = False

# Excel 中文列名到内部英文字段的映射
EXCEL_COLUMN_MAP = {
    '项目类型': 'project_type',
    '项目名称': 'project_name',
    '项目业主': 'project_owner',
    '所属最高发电集团': 'project_owner',
    '装机容量（交流侧）': 'capacity',
    '申报价格': 'declared_price',
    '机制电量比例': 'mechanism_ratio',
    '（承诺）投产时间': 'commission_date',
    '申报电量': 'declared_electricity',
    '机制电量（中标电量）': 'awarded_electricity',
    '机制电价': 'awarded_price',
}


# ================================================================
# 数据对象定义
# ================================================================

@dataclass
class UnitBidRecord:
    """单条竞价记录"""
    year: int
    project_type: str
    project_name: str
    project_owner: str
    capacity: Optional[float]  # MW
    declared_price: Optional[float]  # 元/MWh
    mechanism_ratio: Optional[float]  # %
    commission_date: str
    declared_electricity: Optional[float]  # MWh
    awarded_electricity: Optional[float]  # MWh
    awarded_price: Optional[float]  # 元/MWh
    month: Optional[int] = None  # 月度扩展预留


@dataclass
class UnitBidHistory:
    """机组竞价历史（多年度/多月度数据）"""
    project_name: str
    project_type: str
    project_owner: str
    capacity: Optional[float]
    # 键改为 (year, month) 元组，month 为 None 表示年度数据
    records: Dict[Tuple[int, Optional[int]], UnitBidRecord] = field(default_factory=dict)

    def get_price_history(self) -> List[Tuple[int, float]]:
        """获取历年申报价格 [(year, price), ...]"""
        return sorted([
            (year, r.declared_price)
            for (year, month), r in self.records.items()
            if r.declared_price is not None
        ])

    def get_electricity_history(self) -> List[Tuple[int, float]]:
        """获取历年申报电量 [(year, elec), ...]"""
        return sorted([
            (year, r.declared_electricity)
            for (year, month), r in self.records.items()
            if r.declared_electricity is not None
        ])

    def get_record(self, year: int) -> Optional[UnitBidRecord]:
        """获取指定年度的记录（兼容旧接口，内部使用 (year, None) 键）"""
        return self.records.get((year, None))

    def get_latest_record(self) -> Optional[UnitBidRecord]:
        """获取最新时间记录"""
        if not self.records:
            return None
        latest_key = max(self.records.keys())
        return self.records[latest_key]

    def get_previous_record(self, year: int) -> Optional[UnitBidRecord]:
        """获取指定年度上一年的记录"""
        keys = sorted(self.records.keys())
        target = (year, None)
        idx = keys.index(target) if target in keys else -1
        if idx > 0:
            return self.records[keys[idx - 1]]
        return None


@dataclass
class MarketData:
    """市场数据容器"""
    units: Dict[str, UnitBidHistory] = field(default_factory=dict)  # project_name -> history
    years: List[int] = field(default_factory=list)

    def get_owner_history(self, owner: str) -> Dict[str, UnitBidHistory]:
        """获取同一业主的所有机组历史"""
        return {
            name: hist for name, hist in self.units.items()
            if hist.project_owner == owner
        }

    def get_type_units(self, project_type: str, year: int) -> List[UnitBidRecord]:
        """获取指定类型和年度的所有机组记录"""
        records = []
        for hist in self.units.values():
            if hist.project_type == project_type:
                record = hist.get_record(year)
                if record:
                    records.append(record)
        return records

    def get_all_records_by_year(self, year: int) -> List[UnitBidRecord]:
        """获取指定年度的所有记录"""
        records = []
        for hist in self.units.values():
            record = hist.get_record(year)
            if record:
                records.append(record)
        return records


# ================================================================
# 数据加载函数
# ================================================================

def _parse_float(v) -> Optional[float]:
    """安全地将值转换为浮点数。"""
    if v is None or str(v).strip() == '':
        return None
    try:
        return float(v)
    except (ValueError, TypeError):
        return None


@dataclass
class DataFileInfo:
    """输入文件元数据"""
    path: str
    year: int
    month: Optional[int] = None


def _extract_time_from_filename(file_path: str) -> Tuple[Optional[int], Optional[int]]:
    """
    从文件名中提取年份和月份。
    支持格式：unit-2025.xlsx, unit-2025年.xlsx, unit-2024年06月.xlsx
    返回：(year, month)，若未解析到则返回 (None, None)
    """
    import re
    basename = os.path.basename(file_path)
    basename_no_ext = os.path.splitext(basename)[0]

    # 优先匹配中文"年""月"格式：2024年 或 2024年06月
    pattern = r'(\d{4})年(?:(\d{1,2})月)?'
    match = re.search(pattern, basename_no_ext)
    if match:
        year = int(match.group(1))
        month = int(match.group(2)) if match.group(2) else None
        if 2000 <= year <= 2100:
            return year, month

    # 兜底：匹配纯数字4位年份
    for part in basename_no_ext.replace('-', '_').split('_'):
        if part.isdigit() and len(part) == 4:
            year = int(part)
            if 2000 <= year <= 2100:
                return year, None

    return None, None


def _extract_year_from_filename(file_path: str) -> Optional[int]:
    """从文件名中提取年份（兼容旧接口）。"""
    year, _ = _extract_time_from_filename(file_path)
    return year


def _read_excel_rows(file_path: str) -> List[Dict[str, str]]:
    """使用 pandas 读取 Excel 文件，将中文列名映射为英文字段。"""
    if not PANDAS_AVAILABLE:
        raise ImportError("读取 Excel 需要安装 pandas，请运行: pip install pandas openpyxl")
    df = pd.read_excel(file_path, dtype=str)
    # 去除列名中的空格与换行符，便于匹配
    df.columns = [str(c).strip().replace('\n', '').replace('\r', '') for c in df.columns]
    rows = []
    for _, row in df.iterrows():
        mapped = {}
        for cn, en in EXCEL_COLUMN_MAP.items():
            # 同样对映射键做标准化处理
            clean_cn = cn.replace('\n', '').replace('\r', '')
            if clean_cn in row:
                mapped[en] = str(row[clean_cn]) if pd.notna(row[clean_cn]) else ''
        rows.append(mapped)
    return rows


def load_market_data(data_files: List[DataFileInfo]) -> MarketData:
    """
    加载多年度市场数据（支持 .csv 与 .xlsx）。

    Args:
        data_files: 数据文件信息列表（每个文件代表一年度/一月度数据）

    Returns:
        MarketData对象
    """
    market_data = MarketData()
    all_years = set()

    for file_info in data_files:
        if not os.path.exists(file_info.path):
            continue

        year = file_info.year
        month = file_info.month
        if year is None:
            continue
        all_years.add(year)

        is_excel = file_info.path.lower().endswith(('.xlsx', '.xls'))

        if is_excel:
            rows = _read_excel_rows(file_info.path)
        else:
            with open(file_info.path, 'r', encoding='utf-8-sig') as f:
                rows = list(csv.DictReader(f))

        for row in rows:
            project_name = row.get('project_name', '').strip()
            if not project_name:
                continue

            # 获取或创建机组历史
            if project_name not in market_data.units:
                market_data.units[project_name] = UnitBidHistory(
                    project_name=project_name,
                    project_type=row.get('project_type', ''),
                    project_owner=row.get('project_owner', ''),
                    capacity=_parse_float(row.get('capacity')),
                )

            # 创建记录
            record = UnitBidRecord(
                year=year,
                project_type=row.get('project_type', ''),
                project_name=project_name,
                project_owner=row.get('project_owner', ''),
                capacity=_parse_float(row.get('capacity')),
                declared_price=_parse_float(row.get('declared_price')),
                mechanism_ratio=_parse_float(row.get('mechanism_ratio')),
                commission_date=row.get('commission_date', ''),
                declared_electricity=_parse_float(row.get('declared_electricity')),
                awarded_electricity=_parse_float(row.get('awarded_electricity')),
                awarded_price=_parse_float(row.get('awarded_price')),
                month=month,
            )

            market_data.units[project_name].records[(year, month)] = record

    market_data.years = sorted(all_years)
    return market_data


# ================================================================
# 报价类指标计算
# ================================================================

def calc_bid_level(price_history: List[Tuple[int, float]]) -> Optional[float]:
    """
    报价水平：多年申报价格的算术平均数。

    Args:
        price_history: [(year, price), ...] 历年价格

    Returns:
        平均价格
    """
    prices = [p for _, p in price_history]
    if not prices:
        return None
    return sum(prices) / len(prices)


def calc_bid_position_coeff(
    declared_price: float,
    year_prices: List[float]
) -> Optional[float]:
    """
    报价位置系数：将报价归一化至[0,1]区间。
    使用同一年度所有机组的申报价格最大最小值。

    Args:
        declared_price: 当前机组申报价格
        year_prices: 同年度所有机组的申报价格列表

    Returns:
        [0,1]之间的位置系数
    """
    if declared_price is None or not year_prices:
        return None

    valid_prices = [p for p in year_prices if p is not None]
    if len(valid_prices) < 2:
        return None

    pmin = min(valid_prices)
    pmax = max(valid_prices)

    if pmax <= pmin:
        return None

    coeff = (declared_price - pmin) / (pmax - pmin)
    return max(0.0, min(1.0, coeff))


def calc_bid_dispersion(price_history: List[Tuple[int, float]]) -> Optional[float]:
    """
    报价离散度：多年申报价格的标准差。

    Args:
        price_history: [(year, price), ...] 历年价格

    Returns:
        标准差，单年数据返回0
    """
    prices = [p for _, p in price_history]
    n = len(prices)
    if n == 0:
        return None
    if n == 1:
        return 0.0

    mean_price = sum(prices) / n
    variance = sum((p - mean_price) ** 2 for p in prices) / n
    return math.sqrt(variance)


def calc_bid_change_rate(
    current_price: Optional[float],
    previous_price: Optional[float]
) -> Optional[float]:
    """
    报价变动率：当年与上一年数据的相对变化率。

    Args:
        current_price: 当年申报价格
        previous_price: 上一年申报价格

    Returns:
        变动率
    """
    if current_price is None or previous_price is None or previous_price == 0:
        return None
    return (current_price - previous_price) / previous_price


def minmax_normalize_vectors(vectors: Dict[str, List[float]]) -> Dict[str, List[float]]:
    """
    按维度做min-max归一化。
    对每一维，跨所有向量取最大值和最小值，将该维映射到[0,1]区间。
    若某维最大值等于最小值（所有机组该维相同），该维置0。
    """
    if not vectors:
        return {}
    names = list(vectors.keys())
    dims = len(vectors[names[0]])
    mins = [min(vectors[n][d] for n in names) for d in range(dims)]
    maxs = [max(vectors[n][d] for n in names) for d in range(dims)]
    result = {}
    for n in names:
        result[n] = [
            (v - mins[d]) / (maxs[d] - mins[d]) if maxs[d] > mins[d] else 0.0
            for d, v in enumerate(vectors[n])
        ]
    return result


def calc_euclidean_distance(vec1: List[float], vec2: List[float]) -> float:
    """计算两个向量的欧氏距离。"""
    if len(vec1) != len(vec2) or len(vec1) == 0:
        return 0.0
    return math.sqrt(sum((a - b) ** 2 for a, b in zip(vec1, vec2)))


def calc_bid_consistency_index(
    unit_index: int,
    all_vectors: List[List[float]],
    radius: float = 0.05
) -> Optional[float]:
    """
    报价一致性指数（归一化局部密度）。

    计算步骤：
    1. 量价特征向量已按维度做min-max归一化（每维映射到[0,1]）
    2. 计算当前机组与其他机组特征向量的欧氏距离
    3. 统计欧氏距离不超过邻域半径 radius 的其他机组数量（仅跳过自身）
    4. 用该数量除以总机组数量，得到归一化局部密度

    Args:
        unit_index: 当前机组向量在 all_vectors 中的索引
        all_vectors: 所有机组的量价特征向量列表（已min-max归一化）
        radius: 邻域半径（暂取0.05）

    Returns:
        [0,1]之间的归一化局部密度，值越大表示周边机组越密集
    """
    if not all_vectors or len(all_vectors) < 2:
        return None

    unit_vec = all_vectors[unit_index]

    neighbor_count = 0
    for i, vec in enumerate(all_vectors):
        if i == unit_index:
            continue
        dist = calc_euclidean_distance(unit_vec, vec)
        if dist <= radius:
            neighbor_count += 1

    return neighbor_count / len(all_vectors)


# ================================================================
# 电量类指标计算
# ================================================================

def calc_declared_elec_ratio(
    declared_electricity: float,
    total_declared_electricity: float
) -> Optional[float]:
    """申报电量占比。"""
    if declared_electricity is None or total_declared_electricity is None:
        return None
    if total_declared_electricity == 0:
        return None
    return declared_electricity / total_declared_electricity


def calc_award_rate(
    awarded_electricity: float,
    declared_electricity: float
) -> Optional[float]:
    """机制电量竞得率。"""
    if awarded_electricity is None or declared_electricity is None:
        return None
    if declared_electricity == 0:
        return None
    return awarded_electricity / declared_electricity


def calc_theoretical_electricity(
    capacity: float,
    utilization_hours: float,
    own_use_rate: float,
    mechanism_ratio_cap: float
) -> Optional[float]:
    """
    计算理论最大可申报电量。
    公式: Qi,理论 = Ci × Hi × (1-ηi) × ρmax
    """
    if any(v is None for v in [capacity, utilization_hours, own_use_rate, mechanism_ratio_cap]):
        return None

    our = own_use_rate / 100.0
    ratio_cap = mechanism_ratio_cap / 100.0

    return capacity * utilization_hours * (1 - our) * ratio_cap


def calc_capacity_declared_ratio(
    declared_electricity: float,
    theoretical_electricity: float
) -> Optional[float]:
    """容量申报率。"""
    if declared_electricity is None or theoretical_electricity is None:
        return None
    if theoretical_electricity == 0:
        return None
    return declared_electricity / theoretical_electricity


def identify_marginal_units(
    records: List[UnitBidRecord],
    clearing_price: Optional[float] = None
) -> Tuple[Optional[float], List[str]]:
    """
    识别边际机组。
    规则：
    1. 找申报价格(declared_price) == 出清价格的机组
    2. 如果不存在，找中标电量 < 申报电量的机组

    Args:
        records: 所有机组的竞价记录
        clearing_price: 外部指定的出清价格（可选）

    Returns:
        (边际价格, 边际机组名称列表)
    """
    awarded_records = [
        r for r in records
        if r.awarded_price is not None and r.awarded_electricity is not None and r.awarded_electricity > 0
    ]

    if not awarded_records:
        return None, []

    # 如果没有外部指定出清价格，从数据中计算（中标价格的最大值）
    if clearing_price is None:
        clearing_price = max(r.awarded_price for r in awarded_records)

    # 步骤1：找 declared_price == clearing_price 的机组
    marginal_units = [
        r.project_name for r in awarded_records
        if r.declared_price is not None and abs(r.declared_price - clearing_price) < 1e-6
    ]

    if marginal_units:
        return clearing_price, marginal_units

    # 步骤2：找 awarded_electricity < declared_electricity 的机组
    partial_award_records = [
        r for r in awarded_records
        if r.declared_electricity is not None and r.awarded_electricity < r.declared_electricity
    ]

    if partial_award_records:
        marginal_units = [r.project_name for r in partial_award_records]
        return clearing_price, marginal_units

    return clearing_price, []


def calc_marginal_elec_ratio(
    marginal_declared_elec: float,
    total_marginal_declared_elec: float
) -> Optional[float]:
    """边际电量占比。"""
    if marginal_declared_elec is None or total_marginal_declared_elec is None:
        return None
    if total_marginal_declared_elec == 0:
        return None
    return marginal_declared_elec / total_marginal_declared_elec


# ================================================================
# 电价类指标计算
# ================================================================

def calc_price_deviation(declared_price: float, lcoe: float) -> Optional[float]:
    """报价偏离度。"""
    if declared_price is None or lcoe is None or lcoe == 0:
        return None
    return (declared_price - lcoe) / lcoe


def calc_revenue_deviation(
    actual_revenue: float,
    reasonable_revenue: float
) -> Optional[float]:
    """收益偏离度。"""
    if actual_revenue is None or reasonable_revenue is None:
        return None
    if reasonable_revenue == 0:
        return None
    return (actual_revenue - reasonable_revenue) / reasonable_revenue


# ================================================================
# 收益计算
# ================================================================

def calculate_total_revenue(
    awarded_electricity: float,
    capacity: float,
    utilization_hours: float,
    awarded_price: float,
    spot_price: float,
    green_certificate_price: float,
    mechanism_ratio: float,
    own_use_rate: float
) -> Dict[str, float]:
    """
    计算市场主体总收益（分解为三部分）。

    Args:
        awarded_electricity: 中标机制电量 (MWh)
        capacity: 装机容量 (MW)
        utilization_hours: 年利用小时数 (h)
        awarded_price: 中标电价 (元/MWh)
        spot_price: 现货参考电价 (元/MWh)
        green_certificate_price: 绿证价格 (元/MWh)
        mechanism_ratio: 机制电量比例 (%)
        own_use_rate: 厂用电率 (%)

    Returns:
        {
            'mechanism_revenue': 机制电量收益,
            'spot_revenue': 现货市场收益,
            'gc_revenue': 绿证收益,
            'total_revenue': 总收益
        }
    """
    # 月度机制电量收益 = (中标电量/12) × (中标电价 - 现货参考电价)
    monthly_awarded = awarded_electricity / 12.0
    mechanism_revenue = monthly_awarded * (awarded_price - spot_price)

    # 月度现货市场收益 = (装机容量 × 年利用小时数 / 12) × 现货参考电价
    # 考虑厂用电率：实际发电量 = 装机容量 × 小时数 × (1 - 厂用电率)
    monthly_hours = utilization_hours / 12.0
    effective_capacity = capacity * (1 - own_use_rate / 100.0)
    spot_generation = effective_capacity * monthly_hours
    spot_revenue = spot_generation * spot_price

    # 绿证收益 = (月度总发电量 - 月度机制中标电量) × 绿证价格
    gc_eligible_electricity = max(0, spot_generation - monthly_awarded)
    gc_revenue = gc_eligible_electricity * green_certificate_price

    total_revenue = mechanism_revenue + spot_revenue + gc_revenue

    return {
        'mechanism_revenue': mechanism_revenue,
        'spot_revenue': spot_revenue,
        'gc_revenue': gc_revenue,
        'total_revenue': total_revenue,
    }


def calculate_reasonable_revenue(
    capacity: float,
    utilization_hours: float,
    own_use_rate: float,
    lcoe: float,
    profit_coefficient: float = 1.1
) -> Optional[float]:
    """
    计算理论合理收益。

    Args:
        capacity: 装机容量 (MW)
        utilization_hours: 年利用小时数 (h)
        own_use_rate: 厂用电率 (%)
        lcoe: 度电成本 (元/MWh)
        profit_coefficient: 收益系数（默认1.1表示成本价的1.1倍）

    Returns:
        年合理收益
    """
    if any(v is None for v in [capacity, utilization_hours, own_use_rate, lcoe]):
        return None

    # 年发电量 = 装机容量 × 利用小时数 × (1 - 厂用电率)
    effective_capacity = capacity * (1 - own_use_rate / 100.0)
    annual_generation = effective_capacity * utilization_hours

    # 合理收益单价 = 成本价 × 收益系数
    reasonable_price = lcoe * profit_coefficient

    return annual_generation * reasonable_price


# ================================================================
# 主分析函数
# ================================================================

@dataclass
class AnalysisConfig:
    """分析配置参数"""
    price_cap_max: Optional[float] = None
    price_cap_min: Optional[float] = None
    clearing_price: Optional[float] = None  # 实际出清价格
    theory_clearing_price: Optional[float] = None  # 理论出清价格
    spot_price: Optional[float] = None  # 现货参考价格
    gc_price: Optional[float] = None  # 绿证价格
    type_params: Dict[str, Dict[str, float]] = field(default_factory=dict)
    lcoe_map: Dict[str, float] = field(default_factory=dict)  # 机组类型 -> LCOE
    profit_coefficient: float = 1.1


def analyze_bidding_behavior(
    data_files: List[str],
    config: AnalysisConfig
) -> Dict[str, Any]:
    """
    主分析函数：对每台机组计算所有竞价行为指标。

    Args:
        data_files: 多年度数据文件路径列表
        config: 分析配置参数

    Returns:
        包含各类指标的字典
    """
    # 加载市场数据
    market_data = load_market_data(data_files)
    if not market_data.units:
        return {'error': '无数据或文件不存在'}

    # 确定最新年度
    latest_year = max(market_data.years)
    latest_records = market_data.get_all_records_by_year(latest_year)

    if not latest_records:
        return {'error': f'最新年度 {latest_year} 无数据'}

    # 计算市场级统计
    total_declared_electricity = sum(
        r.declared_electricity or 0 for r in latest_records
    )

    # 获取同一年度所有机组的申报价格（用于报价位置系数）
    year_declared_prices = [r.declared_price for r in latest_records if r.declared_price is not None]

    # 识别边际机组（使用外部指定的出清价格，如果没有则自动计算）
    clearing_price, marginal_unit_names = identify_marginal_units(
        latest_records,
        config.clearing_price
    )

    # 使用外部指定的出清价格优先
    if config.clearing_price is not None:
        clearing_price = config.clearing_price

    # 计算边际档总申报电量
    total_marginal_declared = sum(
        r.declared_electricity or 0
        for r in latest_records
        if r.project_name in marginal_unit_names
    )

    # 构建多年度特征向量（用于一致性指数）
    # 向量格式: [price_y1, ..., price_yN, elec_y1, ..., elec_yN]
    # 缺少某年度数据的机组在该年度位置补0，确保所有向量维数一致
    # min-max归一化：按维度跨所有机组取最大最小值，映射到[0,1]
    raw_vectors = {}
    for project_name, history in market_data.units.items():
        prices = []
        elecs = []
        for year in market_data.years:
            record = history.get_record(year)
            prices.append(record.declared_price if record and record.declared_price is not None else 0.0)
            elecs.append(record.declared_electricity if record and record.declared_electricity is not None else 0.0)
        vector = prices + elecs
        if vector:  # 只有有数据的机组才加入
            raw_vectors[project_name] = vector
    unit_vectors = minmax_normalize_vectors(raw_vectors)

    all_vectors = list(unit_vectors.values())
    unit_vector_idx = {name: i for i, name in enumerate(unit_vectors)}

    # 对每台机组计算指标
    unit_results = {}
    for project_name, history in market_data.units.items():
        latest_record = history.get_latest_record()
        if latest_record is None or latest_record.year != latest_year:
            # 跳过最新年度没有申报的机组
            continue

        # 获取历史数据
        price_history = history.get_price_history()
        prev_record = history.get_previous_record(latest_year)

        # 报价类指标
        bid_level = calc_bid_level(price_history)
        bid_dispersion = calc_bid_dispersion(price_history)
        bid_change = calc_bid_change_rate(
            latest_record.declared_price,
            prev_record.declared_price if prev_record else None
        )
        bid_position = calc_bid_position_coeff(
            latest_record.declared_price,
            year_declared_prices
        )

        # 报价一致性指数（基于多年度向量）
        # 只有拥有多年度数据的机组才能计算
        consistency_index = None
        if project_name in unit_vector_idx and len(unit_vectors[project_name]) >= 4:
            # 至少有2年的数据（每年2个维度）
            consistency_index = calc_bid_consistency_index(
                unit_vector_idx[project_name],
                all_vectors
            )

        # 电量类指标
        declared_elec = latest_record.declared_electricity
        awarded_elec = latest_record.awarded_electricity

        declared_ratio = calc_declared_elec_ratio(
            declared_elec,
            total_declared_electricity
        )
        award_rate = calc_award_rate(awarded_elec, declared_elec)

        # 容量申报率
        capacity_declared_ratio = None
        type_param = config.type_params.get(history.project_type, {})
        if type_param:
            theoretical = calc_theoretical_electricity(
                history.capacity,
                type_param.get('utilization_hours'),
                type_param.get('own_use_rate'),
                type_param.get('mechanism_ratio_cap')
            )
            if theoretical:
                capacity_declared_ratio = calc_capacity_declared_ratio(
                    declared_elec,
                    theoretical
                )

        # 边际电量占比
        is_marginal = project_name in marginal_unit_names
        marginal_ratio = None
        if is_marginal:
            marginal_ratio = calc_marginal_elec_ratio(
                declared_elec,
                total_marginal_declared
            )

        # 报价偏离度
        price_dev = None
        lcoe = config.lcoe_map.get(history.project_type)
        if lcoe is not None and latest_record.declared_price is not None:
            price_dev = calc_price_deviation(latest_record.declared_price, lcoe)

        # 收益计算
        revenue_breakdown = None
        reasonable_revenue = None
        revenue_dev = None

        if all(v is not None for v in [
            awarded_elec, history.capacity, type_param.get('utilization_hours'),
            latest_record.awarded_price, config.spot_price, config.gc_price
        ]):
            revenue_breakdown = calculate_total_revenue(
                awarded_electricity=awarded_elec,
                capacity=history.capacity,
                utilization_hours=type_param.get('utilization_hours'),
                awarded_price=latest_record.awarded_price,
                spot_price=config.spot_price,
                green_certificate_price=config.gc_price,
                mechanism_ratio=latest_record.mechanism_ratio or 0,
                own_use_rate=type_param.get('own_use_rate', 0)
            )

        # 合理收益（除以12转为月度口径，与实际收益一致）
        if lcoe is not None:
            annual_reasonable = calculate_reasonable_revenue(
                capacity=history.capacity,
                utilization_hours=type_param.get('utilization_hours'),
                own_use_rate=type_param.get('own_use_rate', 0),
                lcoe=lcoe,
                profit_coefficient=config.profit_coefficient
            )
            if annual_reasonable is not None:
                reasonable_revenue = annual_reasonable / 12.0

        # 收益偏离度
        if revenue_breakdown and reasonable_revenue:
            revenue_dev = calc_revenue_deviation(
                revenue_breakdown['total_revenue'],
                reasonable_revenue
            )

        unit_results[project_name] = {
            'project_type': history.project_type,
            'project_owner': history.project_owner,
            'year': latest_year,
            'capacity': history.capacity,
            'data_years': [y for y, _ in price_history],
            # 报价类指标
            'bid_level': bid_level,
            'bid_position_coeff': bid_position,
            'bid_dispersion': bid_dispersion,
            'bid_change_rate': bid_change,
            'bid_consistency_index': consistency_index,
            # 电量类指标
            'declared_elec_ratio': declared_ratio,
            'award_rate': award_rate,
            'capacity_declared_ratio': capacity_declared_ratio,
            'marginal_elec_ratio': marginal_ratio,
            'is_marginal': is_marginal,
            # 电价类指标
            'price_deviation': price_dev,
            'revenue_deviation': revenue_dev,
            # 收益分解
            'revenue_breakdown': revenue_breakdown,
            'reasonable_revenue': reasonable_revenue,
        }

    return {
        'by_unit': unit_results,
        'metadata': {
            'data_files': data_files,
            'total_units': len(unit_results),
            'latest_year': latest_year,
            'data_years': market_data.years,
            'config': {
                'price_cap_max': config.price_cap_max,
                'price_cap_min': config.price_cap_min,
                'clearing_price': config.clearing_price,
                'theory_clearing_price': config.theory_clearing_price,
                'spot_price': config.spot_price,
                'gc_price': config.gc_price,
            }
        }
    }


# ================================================================
# 调试入口
# ================================================================

def _format_value(v, precision: int = 4) -> str:
    """格式化数值输出。"""
    if v is None:
        return 'N/A'
    if isinstance(v, float):
        return f'{v:.{precision}f}'
    if isinstance(v, dict):
        return ', '.join([f"{k}: {_format_value(val, precision)}" for k, val in v.items()])
    if isinstance(v, list):
        return str(v)
    return str(v)


def _print_results(results: Dict[str, Any]):
    """打印分析结果到控制台。"""
    print("=" * 100)
    print("竞价行为分析结果（机组级别）")
    print("=" * 100)

    metadata = results.get('metadata', {})
    config = metadata.get('config', {})

    print(f"\n【基本信息】")
    print(f"  数据文件: {metadata.get('data_files', [])}")
    print(f"  机组数量: {metadata.get('total_units', 0)}")
    print(f"  分析年度: {metadata.get('latest_year')}")
    print(f"  数据覆盖年度: {metadata.get('data_years', [])}")

    print(f"\n【配置参数】")
    print(f"  价格上限: {config.get('price_cap_max')} 元/MWh")
    print(f"  价格下限: {config.get('price_cap_min')} 元/MWh")
    print(f"  出清价格: {config.get('clearing_price')} 元/MWh")
    print(f"  理论出清价格: {config.get('theory_clearing_price')} 元/MWh")
    print(f"  现货参考价格: {config.get('spot_price')} 元/MWh")
    print(f"  绿证价格: {config.get('gc_price')} 元/MWh")

    print(f"\n【机组指标详情】")
    print("-" * 100)
    by_unit = results.get('by_unit', {})

    for unit_name, data in sorted(by_unit.items()):
        print(f"\n  机组: {unit_name}")
        print(f"    业主: {data.get('project_owner')}, 类型: {data.get('project_type')}, 年度: {data.get('year')}")
        print(f"    装机容量: {_format_value(data.get('capacity'))} MW")
        print(f"    历史数据年度: {data.get('data_years')}")
        print(f"    --- 报价类指标 ---")
        print(f"    报价水平(多年平均): {_format_value(data.get('bid_level'))}")
        print(f"    报价位置系数: {_format_value(data.get('bid_position_coeff'))}")
        print(f"    报价离散度(多年标准差): {_format_value(data.get('bid_dispersion'))}")
        print(f"    报价变动率(同比): {_format_value(data.get('bid_change_rate'))}")
        print(f"    报价一致性指数: {_format_value(data.get('bid_consistency_index'))}")
        print(f"    --- 电量类指标 ---")
        print(f"    申报电量占比: {_format_value(data.get('declared_elec_ratio'))}")
        print(f"    机制电量竞得率: {_format_value(data.get('award_rate'))}")
        print(f"    容量申报率: {_format_value(data.get('capacity_declared_ratio'))}")
        print(f"    边际电量占比: {_format_value(data.get('marginal_elec_ratio'))}")
        print(f"    是否边际机组: {'是' if data.get('is_marginal') else '否'}")
        print(f"    --- 电价类指标 ---")
        print(f"    报价偏离度: {_format_value(data.get('price_deviation'))}")
        print(f"    收益偏离度: {_format_value(data.get('revenue_deviation'))}")
        print(f"    收益分解: {_format_value(data.get('revenue_breakdown'))}")
        print(f"    理论合理收益: {_format_value(data.get('reasonable_revenue'), 2)}")

    print("\n" + "=" * 100)


def export_results_to_csv(results: Dict[str, Any], output_file: str, market_data: MarketData):
    """将分析结果导出到CSV文件，仅包含指定指标。"""
    # 定义需要导出的指标列（按类别排序）
    columns = [
        'project_name',
        'project_owner',       # 项目业主
        # 报价类指标
        'bid_level',           # 报价水平
        'bid_position_coeff',  # 报价位置系数
        'bid_dispersion',      # 报价离散度
        'bid_change_rate',     # 报价变动率
        'bid_consistency_index', # 报价一致性指数
        # 电量类指标
        'declared_elec_ratio',   # 申报电量占比
        'award_rate',            # 机制电量竞得率
        'capacity_declared_ratio', # 容量申报率
        'marginal_elec_ratio',   # 边际电量占比
        # 电价类指标
        'price_deviation',     # 报价偏离度
        'revenue_deviation',   # 收益偏离度
        # 特征向量
        'bid_curve_vector',    # 报价分布形态特征向量
    ]

    # 先构建所有机组的原始多年度特征向量（顺序：先所有年电价，后所有年电量）
    # 缺少某年度数据的机组在该年度位置补0，确保所有向量维数一致
    raw_vectors = {}
    for unit_name in results.get('by_unit', {}):
        history = market_data.units.get(unit_name)
        if history:
            # 构建向量：[price_2025, price_2026, ..., elec_2025, elec_2026, ...]
            prices = []
            elecs = []
            for year in market_data.years:
                record = history.get_record(year)
                prices.append(record.declared_price if record and record.declared_price is not None else 0.0)
                elecs.append(record.declared_electricity if record and record.declared_electricity is not None else 0.0)
            raw_vectors[unit_name] = prices + elecs
    # min-max归一化：按维度跨所有机组取最大最小值，映射到[0,1]
    norm_vectors = minmax_normalize_vectors(raw_vectors)

    unit_data = []
    for unit_name, data in results.get('by_unit', {}).items():
        row = {'project_name': unit_name}
        # 只提取指定的指标
        for col in columns[1:-1]:  # 跳过 project_name 和 bid_curve_vector
            val = data.get(col)
            # None值用0补位
            row[col] = val if val is not None else 0.0

        vec = norm_vectors.get(unit_name, [])
        row['bid_curve_vector'] = str(vec) if vec else "[]"

        unit_data.append(row)

    if unit_data:
        # 确保输出目录存在
        os.makedirs(os.path.dirname(output_file), exist_ok=True)
        with open(output_file, 'w', encoding='utf-8-sig', newline='') as f:
            writer = csv.DictWriter(f, fieldnames=columns)
            writer.writeheader()
            writer.writerows(unit_data)
        print(f"\n结果已导出到: {output_file}")


def export_results_to_excel(results: Dict[str, Any], output_file: str, market_data: MarketData):
    """将分析结果导出到 Excel 文件，字段与 CSV 版保持一致。"""
    if not PANDAS_AVAILABLE:
        print("警告: pandas 未安装，无法导出 Excel，请运行: pip install pandas openpyxl")
        return

    columns = [
        'project_name',
        'project_owner',
        'bid_level',
        'bid_position_coeff',
        'bid_dispersion',
        'bid_change_rate',
        'bid_consistency_index',
        'declared_elec_ratio',
        'award_rate',
        'capacity_declared_ratio',
        'marginal_elec_ratio',
        'price_deviation',
        'revenue_deviation',
        'bid_curve_vector',
    ]

    # 构建归一化特征向量（与 CSV 版逻辑一致）
    raw_vectors = {}
    for unit_name in results.get('by_unit', {}):
        history = market_data.units.get(unit_name)
        if history:
            prices = []
            elecs = []
            for year in market_data.years:
                record = history.get_record(year)
                prices.append(record.declared_price if record and record.declared_price is not None else 0.0)
                elecs.append(record.declared_electricity if record and record.declared_electricity is not None else 0.0)
            raw_vectors[unit_name] = prices + elecs
    norm_vectors = minmax_normalize_vectors(raw_vectors)

    unit_data = []
    for unit_name, data in results.get('by_unit', {}).items():
        row = {'project_name': unit_name}
        for col in columns[1:-1]:
            val = data.get(col)
            row[col] = val if val is not None else 0.0
        vec = norm_vectors.get(unit_name, [])
        row['bid_curve_vector'] = str(vec) if vec else "[]"
        unit_data.append(row)

    if unit_data:
        os.makedirs(os.path.dirname(output_file), exist_ok=True)
        df = pd.DataFrame(unit_data, columns=columns)

        # 中文列名映射（仅影响 Excel 显示，不改变内部字段名）
        COLUMN_NAME_MAP = {
            'project_name': '项目名称',
            'project_owner': '所属发电集团',
            'bid_level': '报价水平',
            'bid_position_coeff': '报价位置系数',
            'bid_dispersion': '报价离散度',
            'bid_change_rate': '报价变动率',
            'bid_consistency_index': '报价一致性指数',
            'declared_elec_ratio': '申报电量占比',
            'award_rate': '机制电量竞得率',
            'capacity_declared_ratio': '容量申报率',
            'marginal_elec_ratio': '边际电量占比',
            'price_deviation': '报价偏离度',
            'revenue_deviation': '收益偏离',
            'bid_curve_vector': '报价分布形态特征向量',
        }
        df = df.rename(columns=COLUMN_NAME_MAP)

        df.to_excel(output_file, index=False)
        print(f"\n结果已导出到: {output_file}")


if __name__ == "__main__":
    import sys

    # 默认数据文件路径
    root_dir = os.path.dirname(os.path.dirname(__file__))
    default_paths = [
        os.path.join(root_dir, 'Data', 'Input', 'Previous', 'unit-2024年.xlsx'),
        os.path.join(root_dir, 'Data', 'Input', 'Previous', 'unit-2025年.xlsx'),
        os.path.join(root_dir, 'Data', 'Input', 'Current', 'unit-2026年.xlsx'),
    ]

    # 支持命令行参数指定文件
    if len(sys.argv) > 1 and not sys.argv[1].startswith('--'):
        arg_paths = []
        for arg in sys.argv[1:]:
            if arg.startswith('--'):
                break
            arg_paths.append(arg)
    else:
        arg_paths = default_paths

    # 构造 DataFileInfo 并过滤存在的文件
    data_files = []
    for p in arg_paths:
        if os.path.exists(p):
            year, month = _extract_time_from_filename(p)
            if year is not None:
                data_files.append(DataFileInfo(path=p, year=year, month=month))

    if not data_files:
        print("错误: 未找到数据文件")
        sys.exit(1)

    print(f"加载数据文件: {[f.path for f in data_files]}")

    # 创建配置
    config = AnalysisConfig(
        # 价格上下限（可选）
        price_cap_max=350.0,
        price_cap_min=100.0,
        # 出清价格（外部指定，None则自动计算）
        clearing_price=230.0,
        # 理论出清价格（用于计算偏离度）
        theory_clearing_price=210.0,
        # 现货参考价格（用于收益计算）
        spot_price=200.0,
        # 绿证价格（元/MWh）
        gc_price=5.0,
        # 类型参数
        type_params={
            '集中式陆上风电': {
                'utilization_hours': 2417,
                'own_use_rate': 2.27,
                'mechanism_ratio_cap': 70,
            },
            '近海风电': {
                'utilization_hours': 2860,
                'own_use_rate': 2.65,
                'mechanism_ratio_cap': 70,
            },
            '分散式陆上风电': {
                'utilization_hours': 2417,
                'own_use_rate': 2.27,
                'mechanism_ratio_cap': 70,
            },
            '深远海风电': {
                'utilization_hours': 3193,
                'own_use_rate': 3.27,
                'mechanism_ratio_cap': 70,
            },
            '集中式光伏': {
                'utilization_hours': 1253,
                'own_use_rate': 1.59,
                'mechanism_ratio_cap': 80,
            },
            '分散式光伏': {
                'utilization_hours': 1253,
                'own_use_rate': 1.59,
                'mechanism_ratio_cap': 80,
            },
        },
        # 度电成本（按机组类型，单位：元/MWh）
        lcoe_map={
            '集中式陆上风电': 220.0,
            '近海风电': 320.0,
            '分散式陆上风电': 220.0,
            '深远海风电': 380.0,
            '集中式光伏': 200.0,
            '分散式光伏': 220.0,
        },
        # 收益系数
        profit_coefficient=1.1,
    )

    # 运行分析
    results = analyze_bidding_behavior(data_files, config)

    if 'error' in results:
        print(f"错误: {results['error']}")
        sys.exit(1)

    # 打印结果
    _print_results(results)

    # 导出结果到Excel
    output_dir = os.path.join(root_dir, 'Data', 'Output')
    output_file = os.path.join(output_dir, 'bidding_analysis_results.xlsx')
    # 重新加载market_data用于导出特征向量
    market_data = load_market_data(data_files)
    export_results_to_excel(results, output_file, market_data)
