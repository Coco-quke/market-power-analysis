"""
主函数：串联竞价行为分析、层次聚类与异常检测，输出四份 Excel 结果。
"""
import os
import sys

# 将项目根目录加入路径，确保可导入同级模块
project_root = os.path.dirname(os.path.abspath(__file__))
if project_root not in sys.path:
    sys.path.insert(0, project_root)

import pandas as pd

from bidding_behavior import (
    load_market_data,
    analyze_bidding_behavior,
    export_results_to_excel,
    AnalysisConfig,
    minmax_normalize_vectors,
    DataFileInfo,
    _extract_time_from_filename,
)
from abnormal_detection import (
    hierarchical_clustering,
    run_abnormal_detection,
    print_clustering_results,
    print_detection_results,
    export_clustering_to_excel,
    ARCHETYPE_PRICE_ALLIANCE,
    ARCHETYPE_AGGRESSIVE_LOW,
    ARCHETYPE_STRATEGIC_PROBE,
    ARCHETYPE_COST_ORIENTED,
)


def _build_ad_records(results, market_data):
    """
    将 bidding_behavior 的分析结果转换为 abnormal_detection 所需的记录列表。
    同时构造 bid_curve_vector（归一化后的多年度特征向量）。
    """
    # 复用 export_results_to_excel 中的向量构造逻辑
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

    records = []
    for unit_name, data in results.get('by_unit', {}).items():
        row = {
            'project_name': unit_name,
            'project_owner': data.get('project_owner', ''),
            'bid_level': data.get('bid_level') if data.get('bid_level') is not None else 0.0,
            'bid_position_coeff': data.get('bid_position_coeff') if data.get('bid_position_coeff') is not None else 0.0,
            'bid_dispersion': data.get('bid_dispersion') if data.get('bid_dispersion') is not None else 0.0,
            'bid_change_rate': data.get('bid_change_rate') if data.get('bid_change_rate') is not None else 0.0,
            'bid_consistency_index': data.get('bid_consistency_index') if data.get('bid_consistency_index') is not None else 0.0,
            'declared_elec_ratio': data.get('declared_elec_ratio') if data.get('declared_elec_ratio') is not None else 0.0,
            'award_rate': data.get('award_rate') if data.get('award_rate') is not None else 0.0,
            'capacity_declared_ratio': data.get('capacity_declared_ratio') if data.get('capacity_declared_ratio') is not None else 0.0,
            'marginal_elec_ratio': data.get('marginal_elec_ratio') if data.get('marginal_elec_ratio') is not None else 0.0,
            'price_deviation': data.get('price_deviation') if data.get('price_deviation') is not None else 0.0,
            'revenue_deviation': data.get('revenue_deviation') if data.get('revenue_deviation') is not None else 0.0,
            'bid_curve_vector': str(norm_vectors.get(unit_name, [])),
        }
        records.append(row)
    return records


def run_analysis(
    previous_dir: str,
    current_dir: str,
    output_dir: str,
    progress_callback=None,
    log_callback=None
):
    """
    核心分析逻辑，供GUI调用。

    Args:
        previous_dir: 历史数据文件夹路径
        current_dir: 最新数据文件夹路径
        output_dir: 输出目录路径
        progress_callback: 进度回调函数，接收0-100的整数
        log_callback: 日志回调函数，接收字符串

    Returns:
        包含异常主体DataFrame和输出文件路径的字典
    """
    def _log(msg):
        if log_callback:
            log_callback(msg)
        else:
            print(msg)

    def _progress(val):
        if progress_callback:
            progress_callback(val)

    # ------------------------------------------------------------------
    # 1. 路径定义：扫描文件夹中的Excel文件
    # ------------------------------------------------------------------
    import glob
    data_files = []

    for folder, label in [(previous_dir, '历史'), (current_dir, '最新')]:
        if not os.path.exists(folder):
            raise FileNotFoundError(f"{label}数据文件夹不存在: {folder}")
        excel_files = glob.glob(os.path.join(folder, '*.xlsx')) + glob.glob(os.path.join(folder, '*.xls'))
        for p in sorted(excel_files):
            year, month = _extract_time_from_filename(p)
            if year is not None:
                data_files.append(DataFileInfo(path=p, year=year, month=month))
                _log(f"发现{label}数据: {os.path.basename(p)} ({year}年)")

    if not data_files:
        raise FileNotFoundError("未找到输入数据文件，请确认文件夹中包含.xlsx文件")

    os.makedirs(output_dir, exist_ok=True)

    output_indicator = os.path.join(output_dir, '指标计算值.xlsx')
    output_abnormal = os.path.join(output_dir, '异常行为主体.xlsx')
    output_normal = os.path.join(output_dir, '正常主体名单.xlsx')
    output_cluster = os.path.join(output_dir, '聚类结果.xlsx')

    # ------------------------------------------------------------------
    # 2. 配置创建
    # ------------------------------------------------------------------
    config = AnalysisConfig(
        price_cap_max=350.0,
        price_cap_min=100.0,
        clearing_price=230.0,
        theory_clearing_price=210.0,
        spot_price=200.0,
        gc_price=5.0,
        type_params={
            '集中式陆上风电': {'utilization_hours': 2417, 'own_use_rate': 2.27, 'mechanism_ratio_cap': 70},
            '近海风电': {'utilization_hours': 2860, 'own_use_rate': 2.65, 'mechanism_ratio_cap': 70},
            '分散式陆上风电': {'utilization_hours': 2417, 'own_use_rate': 2.27, 'mechanism_ratio_cap': 70},
            '深远海风电': {'utilization_hours': 3193, 'own_use_rate': 3.27, 'mechanism_ratio_cap': 70},
            '集中式光伏': {'utilization_hours': 1253, 'own_use_rate': 1.59, 'mechanism_ratio_cap': 80},
            '分散式光伏': {'utilization_hours': 1253, 'own_use_rate': 1.59, 'mechanism_ratio_cap': 80},
        },
        lcoe_map={
            '集中式陆上风电': 220.0,
            '近海风电': 320.0,
            '分散式陆上风电': 220.0,
            '深远海风电': 380.0,
            '集中式光伏': 200.0,
            '分散式光伏': 220.0,
        },
        profit_coefficient=1.1,
    )

    # ------------------------------------------------------------------
    # 3. 模块A：竞价行为分析
    # ------------------------------------------------------------------
    _progress(10)
    _log("=" * 60)
    _log("【模块A】竞价行为分析")
    _log("=" * 60)
    results = analyze_bidding_behavior(data_files, config)
    if 'error' in results:
        raise RuntimeError(results['error'])

    market_data = load_market_data(data_files)
    export_results_to_excel(results, output_indicator, market_data)
    _log(f"指标计算值已导出: {output_indicator}")
    _progress(30)

    # ------------------------------------------------------------------
    # 4. 构造 abnormal_detection 所需的记录列表
    # ------------------------------------------------------------------
    ad_records = _build_ad_records(results, market_data)
    record_map = {r['project_name']: r for r in ad_records}

    # ------------------------------------------------------------------
    # 5. 模块B-1：层次聚类
    # ------------------------------------------------------------------
    _progress(40)
    _log("\n" + "=" * 60)
    _log("【模块B-1】层次聚类")
    _log("=" * 60)
    cluster_results = hierarchical_clustering(ad_records)
    if 'error' in cluster_results:
        raise RuntimeError(cluster_results['error'])
    print_clustering_results(cluster_results)
    export_clustering_to_excel(cluster_results, ad_records, output_cluster)
    _log(f"聚类结果已导出: {output_cluster}")
    _progress(60)

    # ------------------------------------------------------------------
    # 6. 按聚类结果预分类
    # ------------------------------------------------------------------
    price_alliance_set = set()
    aggressive_low_set = set()
    pending_set = set()

    for label, members in cluster_results['clusters'].items():
        archetype = cluster_results['archetype_map'][label]
        if archetype == ARCHETYPE_PRICE_ALLIANCE:
            price_alliance_set.update(members)
        elif archetype == ARCHETYPE_AGGRESSIVE_LOW:
            aggressive_low_set.update(members)
        else:
            pending_set.update(members)

    _log(f"\n聚类预分类: 价格联盟型 {len(price_alliance_set)} 个, 待检测 {len(pending_set)} 个")

    # ------------------------------------------------------------------
    # 7. 对全量主体执行异常检测
    # ------------------------------------------------------------------
    _progress(70)
    _log("\n" + "=" * 60)
    _log("【模块B-2】联合异常检测")
    _log("=" * 60)
    detection_results = run_abnormal_detection(
        ad_records,
        feature_cols=['bid_position_coeff', 'price_deviation', 'revenue_deviation']
    )
    if 'error' in detection_results:
        raise RuntimeError(detection_results['error'])
    print_detection_results(detection_results)

    # ------------------------------------------------------------------
    # 8. 按规则重新分配
    # ------------------------------------------------------------------
    normal_set = set()

    for unit in detection_results['classification']['both_abnormal']:
        if unit in pending_set:
            record = record_map.get(unit)
            if record and record['price_deviation'] < 0:
                aggressive_low_set.add(unit)
            else:
                normal_set.add(unit)

    for unit in detection_results['classification']['rf_only_abnormal']:
        if unit in pending_set:
            normal_set.add(unit)

    for unit in detection_results['classification']['threshold_only_abnormal']:
        if unit in pending_set:
            normal_set.add(unit)

    for unit in detection_results['classification']['both_normal']:
        if unit in pending_set:
            normal_set.add(unit)

    _log(f"\n最终分类: 价格联盟型 {len(price_alliance_set)} 个, 激进低价型 {len(aggressive_low_set)} 个, 正常主体 {len(normal_set)} 个")
    _progress(85)

    # ------------------------------------------------------------------
    # 9. 构造异常主体DataFrame
    # ------------------------------------------------------------------
    abnormal_rows = []
    for unit in sorted(price_alliance_set):
        r = record_map[unit]
        abnormal_rows.append({
            '市场主体名称': unit,
            '所属发电集团': r['project_owner'],
            '行为类型': '价格联盟型',
            '报价偏离度': round(r['price_deviation'], 4),
            '收益偏离度': round(r['revenue_deviation'], 4),
        })
    for unit in sorted(aggressive_low_set):
        r = record_map[unit]
        abnormal_rows.append({
            '市场主体名称': unit,
            '所属发电集团': r['project_owner'],
            '行为类型': '激进低价型',
            '报价偏离度': round(r['price_deviation'], 4),
            '收益偏离度': round(r['revenue_deviation'], 4),
        })

    df_abnormal = pd.DataFrame(abnormal_rows) if abnormal_rows else pd.DataFrame(
        columns=['市场主体名称', '所属发电集团', '行为类型', '报价偏离度', '收益偏离度']
    )
    df_abnormal.to_excel(output_abnormal, index=False)
    _log(f"异常行为主体已导出: {output_abnormal}")

    # ------------------------------------------------------------------
    # 10. 导出正常主体名单
    # ------------------------------------------------------------------
    normal_rows = []
    for unit in sorted(normal_set):
        r = record_map[unit]
        normal_rows.append({
            '市场主体名称': unit,
            '所属发电集团': r['project_owner'],
            '报价偏离度': round(r['price_deviation'], 4),
            '收益偏离度': round(r['revenue_deviation'], 4),
        })

    if normal_rows:
        df_normal = pd.DataFrame(normal_rows)
        df_normal.to_excel(output_normal, index=False)
        _log(f"正常主体名单已导出: {output_normal}")

    _progress(100)
    _log("\n" + "=" * 60)
    _log("全部计算完成！")
    _log("=" * 60)

    return {
        'df_abnormal': df_abnormal,
        'output_files': {
            'indicator': output_indicator,
            'abnormal': output_abnormal,
            'normal': output_normal,
            'cluster': output_cluster,
        }
    }


def main():
    # ------------------------------------------------------------------
    # 1. 路径定义
    # ------------------------------------------------------------------
    root_dir = os.path.dirname(os.path.abspath(__file__))
    previous_dir = os.path.join(root_dir, 'Data', 'Input', 'Previous')
    current_dir = os.path.join(root_dir, 'Data', 'Input', 'Current')
    output_dir = os.path.join(root_dir, 'Data', 'Output')

    try:
        result = run_analysis(previous_dir, current_dir, output_dir)
    except Exception as e:
        print(f"错误: {e}")
        sys.exit(1)


if __name__ == '__main__':
    main()
