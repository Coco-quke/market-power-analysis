"""
异常市场主体识别模块
采用随机森林和硬阈值两种方式识别异常市场主体。
"""
import csv
import os
from typing import Dict, List, Optional, Any, Tuple
import ast

# 尝试导入sklearn，如果未安装则给出提示
try:
    from sklearn.ensemble import IsolationForest
    from sklearn.decomposition import PCA
    from sklearn.preprocessing import StandardScaler
    SKLEARN_AVAILABLE = True
except ImportError:
    SKLEARN_AVAILABLE = False
    print("警告: scikit-learn 未安装，请运行: pip install scikit-learn")

# 尝试导入 pandas 以支持 Excel 导出
try:
    import pandas as pd
    PANDAS_AVAILABLE = True
except ImportError:
    PANDAS_AVAILABLE = False


def load_analysis_data(data_file: str) -> List[Dict[str, Any]]:
    """
    加载竞价行为分析结果数据。

    Args:
        data_file: CSV文件路径

    Returns:
        数据记录列表
    """
    if not os.path.exists(data_file):
        return []

    records = []
    with open(data_file, 'r', encoding='utf-8-sig') as f:
        reader = csv.DictReader(f)
        for row in reader:
            record = {
                'project_name': row.get('project_name', ''),
                'project_owner': row.get('project_owner', ''),
                'bid_level': float(row.get('bid_level', 0)) if row.get('bid_level') else 0.0,
                'bid_position_coeff': float(row.get('bid_position_coeff', 0)) if row.get('bid_position_coeff') else 0.0,
                'bid_dispersion': float(row.get('bid_dispersion', 0)) if row.get('bid_dispersion') else 0.0,
                'bid_change_rate': float(row.get('bid_change_rate', 0)) if row.get('bid_change_rate') else 0.0,
                'bid_consistency_index': float(row.get('bid_consistency_index', 0)) if row.get('bid_consistency_index') else 0.0,
                'declared_elec_ratio': float(row.get('declared_elec_ratio', 0)) if row.get('declared_elec_ratio') else 0.0,
                'award_rate': float(row.get('award_rate', 0)) if row.get('award_rate') else 0.0,
                'capacity_declared_ratio': float(row.get('capacity_declared_ratio', 0)) if row.get('capacity_declared_ratio') else 0.0,
                'marginal_elec_ratio': float(row.get('marginal_elec_ratio', 0)) if row.get('marginal_elec_ratio') else 0.0,
                'price_deviation': float(row.get('price_deviation', 0)) if row.get('price_deviation') else 0.0,
                'revenue_deviation': float(row.get('revenue_deviation', 0)) if row.get('revenue_deviation') else 0.0,
                'bid_curve_vector': row.get('bid_curve_vector', '[]'),  # 读取特征向量
            }
            records.append(record)
    return records


def compress_bid_curve_pca(
    records: List[Dict[str, Any]],
    n_components: int = 1
) -> Tuple[Dict[str, List[float]], float]:
    """
    对所有机组的 bid_curve_vector 做PCA降维。

    Args:
        records: 数据记录列表
        n_components: 保留的主成分数量（默认1）

    Returns:
        (project_name -> [pc1, ...] 的字典, 解释方差比例)
    """
    if not SKLEARN_AVAILABLE:
        return {}, 0.0

    # 解析所有 bid_curve_vector
    vectors = []
    names = []
    for record in records:
        vec_str = record.get('bid_curve_vector', '[]')
        try:
            vec = ast.literal_eval(vec_str)
            if vec:
                vectors.append(vec)
                names.append(record['project_name'])
        except:
            pass

    if len(vectors) < 2 or n_components < 1:
        return {}, 0.0

    # 标准化 + PCA
    X = StandardScaler().fit_transform(vectors)
    pca = PCA(n_components=n_components)
    X_pca = pca.fit_transform(X)

    result = {}
    for i, name in enumerate(names):
        result[name] = X_pca[i].tolist()

    explained = float(pca.explained_variance_ratio_[0]) if len(pca.explained_variance_ratio_) > 0 else 0.0
    return result, explained


def extract_features(
    records: List[Dict[str, Any]],
    feature_cols: Optional[List[str]] = None,
    compress_vector: bool = False,
    n_components: int = 1
) -> Tuple[List[List[float]], List[str], List[str]]:
    """
    从记录中提取特征向量和机组名称。

    Args:
        records: 数据记录列表
        feature_cols: 若指定，仅提取对应的基础指标列，跳过向量展开和one-hot编码
        compress_vector: 若True，对bid_curve_vector做PCA降维后再使用
        n_components: PCA保留的主成分数量

    Returns:
        (特征矩阵, 机组名称列表, 特征名称列表)
    """
    features = []
    unit_names = []

    if feature_cols is not None:
        # 简化模式：仅提取指定列
        all_feature_names = list(feature_cols)
        for record in records:
            unit_names.append(record['project_name'])
            feature_vec = [record.get(col, 0.0) for col in feature_cols]
            features.append(feature_vec)
        return features, unit_names, all_feature_names

    # 若启用PCA压缩，先计算主成分并注入记录
    pca_map = {}
    explained = 0.0
    if compress_vector:
        pca_map, explained = compress_bid_curve_pca(records, n_components)
        for record in records:
            name = record['project_name']
            if name in pca_map:
                for i, val in enumerate(pca_map[name]):
                    record[f'pca_vec_{i}'] = val

    # 基础特征列（排除 project_name 和 bid_curve_vector）
    base_feature_cols = [
        'bid_position_coeff',
        'bid_change_rate',
        'bid_consistency_index',
        'price_deviation',
        'revenue_deviation',
        'bid_dispersion',
        'declared_elec_ratio',
        'marginal_elec_ratio',
    ]

    # 确定向量特征名称（PCA压缩或原始展开）
    if compress_vector:
        vector_feature_names = [f'pca_vec_{i}' for i in range(n_components)]
    else:
        vector_feature_names = []
        max_vector_len = 0
        for record in records:
            vec_str = record.get('bid_curve_vector', '[]')
            try:
                vec = ast.literal_eval(vec_str)
                max_vector_len = max(max_vector_len, len(vec))
            except:
                pass
        for i in range(max_vector_len):
            vector_feature_names.append(f'bid_curve_vec_{i}')

    # project_owner one-hot 编码
    owners = sorted(set(r.get('project_owner', '') for r in records if r.get('project_owner')))
    owner_feature_names = [f'owner_{o}' for o in owners]

    # 所有特征名称
    all_feature_names = base_feature_cols + vector_feature_names + owner_feature_names

    for record in records:
        unit_names.append(record['project_name'])

        # 基础特征
        feature_vec = [record[col] for col in base_feature_cols]

        if compress_vector:
            # 使用PCA压缩后的主成分
            for i in range(n_components):
                feature_vec.append(record.get(f'pca_vec_{i}', 0.0))
        else:
            # 展开 bid_curve_vector
            vec_str = record.get('bid_curve_vector', '[]')
            try:
                vec = ast.literal_eval(vec_str)
                while len(vec) < max_vector_len:
                    vec.append(0.0)
                vec = vec[:max_vector_len]
                feature_vec.extend(vec)
            except:
                feature_vec.extend([0.0] * max_vector_len)

        # project_owner one-hot 编码
        owner = record.get('project_owner', '')
        feature_vec.extend([1.0 if owner == o else 0.0 for o in owners])

        features.append(feature_vec)

    return features, unit_names, all_feature_names


def random_forest_detection(
    features: List[List[float]],
    unit_names: List[str],
    contamination: float = 0.1
) -> Dict[str, bool]:
    """
    使用随机森林(IsolationForest)检测异常市场主体。

    Args:
        features: 特征矩阵
        unit_names: 机组名称列表
        contamination: 预期异常样本比例（默认10%），控制判定灵敏度，
                       数值越小判定越严格

    Returns:
        机组名称 -> 是否异常的字典
    """
    if not SKLEARN_AVAILABLE:
        print("错误: scikit-learn 未安装，无法进行随机森林检测")
        return {name: False for name in unit_names}

    if len(features) < 2:
        print("警告: 数据量过少，无法进行随机森林检测")
        return {name: False for name in unit_names}

    # 使用IsolationForest进行异常检测
    # contamination 为预期异常比例，决定判定阈值位置
    clf = IsolationForest(
        n_estimators=100,
        max_samples='auto',
        contamination=contamination,
        random_state=42,
        n_jobs=-1
    )

    # 拟合数据并预测
    clf.fit(features)
    predictions = clf.predict(features)

    # predictions: 1表示正常，-1表示异常
    results = {}
    for name, pred in zip(unit_names, predictions):
        results[name] = (pred == -1)  # -1为异常

    return results


def threshold_detection(records: List[Dict[str, Any]]) -> Dict[str, bool]:
    """
    使用硬阈值方式检测异常市场主体。

    阈值规则:
    - 报价偏离度 < -30% 或 > 30% 为异常
    - 收益偏离度 < -10% 或 > 30% 为异常
    - 任一指标异常则认为主体异常

    Args:
        records: 数据记录列表

    Returns:
        机组名称 -> 是否异常的字典
    """
    results = {}

    for record in records:
        name = record['project_name']
        price_dev = record['price_deviation']
        revenue_dev = record['revenue_deviation']

        # 判断报价偏离度是否异常
        price_abnormal = (price_dev < -0.30) or (price_dev > 0.30)

        # 判断收益偏离度是否异常
        revenue_abnormal = (revenue_dev < -0.10) or (revenue_dev > 0.30)

        # 任一指标异常则认为主体异常
        is_abnormal = price_abnormal or revenue_abnormal

        results[name] = is_abnormal

    return results


def classify_results(
    rf_results: Dict[str, bool],
    threshold_results: Dict[str, bool]
) -> Dict[str, List[str]]:
    """
    根据两种方法的结果将市场主体分为四类。

    Args:
        rf_results: 随机森林检测结果
        threshold_results: 硬阈值检测结果

    Returns:
        四类结果字典
    """
    both_abnormal = []      # 两类都异常
    rf_only_abnormal = []   # 仅随机森林异常
    threshold_only_abnormal = []  # 仅硬阈值异常
    both_normal = []        # 两类都正常

    all_units = set(rf_results.keys()) | set(threshold_results.keys())

    for unit in all_units:
        rf_abnormal = rf_results.get(unit, False)
        th_abnormal = threshold_results.get(unit, False)

        if rf_abnormal and th_abnormal:
            both_abnormal.append(unit)
        elif rf_abnormal and not th_abnormal:
            rf_only_abnormal.append(unit)
        elif not rf_abnormal and th_abnormal:
            threshold_only_abnormal.append(unit)
        else:
            both_normal.append(unit)

    return {
        'both_abnormal': both_abnormal,
        'rf_only_abnormal': rf_only_abnormal,
        'threshold_only_abnormal': threshold_only_abnormal,
        'both_normal': both_normal,
    }


def run_abnormal_detection(
    records: List[Dict[str, Any]],
    contamination: float = 0.1,
    feature_cols: Optional[List[str]] = None
) -> Dict[str, Any]:
    """
    对已加载的记录执行随机森林+硬阈值异常检测。

    Args:
        records: 数据记录列表
        contamination: 随机森林的异常比例估计（默认10%）
        feature_cols: 若指定，仅提取对应列作为随机森林特征（硬阈值不受影响）

    Returns:
        包含检测结果的字典
    """
    if not records:
        return {'error': '无输入记录'}

    print(f"\n参与异常识别的主体数: {len(records)}")

    # 提取特征
    features, unit_names, feature_names = extract_features(records, feature_cols)
    print(f"特征维度: {len(feature_names)} 维")
    print(f"特征列表: {feature_names}")

    # 方法1：随机森林检测
    print(f"\n正在进行随机森林异常检测（contamination={contamination}）...")
    rf_results = random_forest_detection(features, unit_names, contamination)
    rf_abnormal_count = sum(1 for v in rf_results.values() if v)
    print(f"随机森林检测到 {rf_abnormal_count} 个异常主体")

    # 方法2：硬阈值检测
    print("\n正在进行硬阈值异常检测...")
    print("阈值规则:")
    print("  - 报价偏离度 < -30% 或 > 30% 为异常")
    print("  - 收益偏离度 < -10% 或 > 30% 为异常")
    threshold_results = threshold_detection(records)
    th_abnormal_count = sum(1 for v in threshold_results.values() if v)
    print(f"硬阈值检测到 {th_abnormal_count} 个异常主体")

    # 分类结果
    classification = classify_results(rf_results, threshold_results)

    return {
        'total_units': len(records),
        'rf_results': rf_results,
        'threshold_results': threshold_results,
        'classification': classification,
        'records': records,
    }


def detect_abnormal_market_entities(
    data_file: str,
    contamination: float = 0.1
) -> Dict[str, Any]:
    """
    主检测函数：综合使用随机森林和硬阈值方法识别异常市场主体。

    Args:
        data_file: 输入数据文件路径
        contamination: 随机森林的异常比例估计（默认10%）

    Returns:
        包含检测结果的字典
    """
    # 加载数据
    records = load_analysis_data(data_file)
    if not records:
        return {'error': '未找到数据或文件不存在'}

    print(f"加载了 {len(records)} 条记录")
    return run_abnormal_detection(records, contamination)


def print_detection_results(results: Dict[str, Any]):
    """打印检测结果到控制台。"""
    if 'error' in results:
        print(f"错误: {results['error']}")
        return

    print("\n" + "=" * 80)
    print("异常市场主体识别结果")
    print("=" * 80)

    total = results['total_units']
    classification = results['classification']

    print(f"\n【总体统计】")
    print(f"  总市场主体数: {total}")
    print(f"  两类都异常: {len(classification['both_abnormal'])} ({len(classification['both_abnormal'])/total*100:.1f}%)")
    print(f"  仅随机森林异常: {len(classification['rf_only_abnormal'])} ({len(classification['rf_only_abnormal'])/total*100:.1f}%)")
    print(f"  仅硬阈值异常: {len(classification['threshold_only_abnormal'])} ({len(classification['threshold_only_abnormal'])/total*100:.1f}%)")
    print(f"  两类都正常: {len(classification['both_normal'])} ({len(classification['both_normal'])/total*100:.1f}%)")

    # 详细结果
    print(f"\n【第一类：两类方法都异常的市场主体】")
    if classification['both_abnormal']:
        for unit in classification['both_abnormal']:
            record = next((r for r in results['records'] if r['project_name'] == unit), None)
            if record:
                print(f"  - {unit}")
                print(f"    报价偏离度: {record['price_deviation']:.4f} ({record['price_deviation']*100:.2f}%)")
                print(f"    收益偏离度: {record['revenue_deviation']:.4f} ({record['revenue_deviation']*100:.2f}%)")
    else:
        print("  无")

    print(f"\n【第二类：仅随机森林判断异常】")
    if classification['rf_only_abnormal']:
        for unit in classification['rf_only_abnormal']:
            print(f"  - {unit}")
    else:
        print("  无")

    print(f"\n【第三类：仅硬阈值判断异常】")
    if classification['threshold_only_abnormal']:
        for unit in classification['threshold_only_abnormal']:
            record = next((r for r in results['records'] if r['project_name'] == unit), None)
            if record:
                print(f"  - {unit}")
                print(f"    报价偏离度: {record['price_deviation']:.4f} ({record['price_deviation']*100:.2f}%)")
                print(f"    收益偏离度: {record['revenue_deviation']:.4f} ({record['revenue_deviation']*100:.2f}%)")
    else:
        print("  无")

    print(f"\n【第四类：两类方法都正常】")
    if classification['both_normal']:
        for unit in classification['both_normal']:
            record = next((r for r in results['records'] if r['project_name'] == unit), None)
            if record:
                print(f"  - {unit}")
                print(f"    报价偏离度: {record['price_deviation']:.4f} ({record['price_deviation']*100:.2f}%)")
                print(f"    收益偏离度: {record['revenue_deviation']:.4f} ({record['revenue_deviation']*100:.2f}%)")
    else:
        print("  无")

    print("\n" + "=" * 80)


def export_clustering_to_excel(cluster_results: Dict[str, Any],
                               records: List[Dict[str, Any]],
                               output_file: str):
    """将层次聚类结果导出到 Excel。"""
    if not PANDAS_AVAILABLE:
        print("警告: pandas 未安装，无法导出聚类 Excel")
        return

    record_map = {r['project_name']: r for r in records}
    clusters = cluster_results['clusters']
    display_name_map = cluster_results['display_name_map']
    profiles = cluster_results['profiles']

    rows = []
    for label in sorted(clusters.keys()):
        members = clusters[label]
        archetype = display_name_map[label]
        profile = profiles[label]
        for m in sorted(members):
            owner = record_map.get(m, {}).get('project_owner', '')
            rows.append({
                '策略行为类型': archetype,
                '簇号': label,
                '市场主体名称': m,
                '所属发电集团': owner,
                '簇_报价位置系数': round(profile['bid_position_coeff'], 4),
                '簇_报价变动率(绝对值)': round(profile['abs_bid_change_rate'], 4),
                '簇_报价一致性指数': round(profile['bid_consistency_index'], 4),
                '簇_报价偏离度': round(profile['price_deviation'], 4),
                '簇_收益偏离度': round(profile['revenue_deviation'], 4),
            })

    if rows:
        os.makedirs(os.path.dirname(output_file), exist_ok=True)
        df = pd.DataFrame(rows)
        df.to_excel(output_file, index=False)
        print(f"聚类结果已导出到: {output_file}")


def export_abnormal_to_excel(results: Dict[str, Any], output_file: str):
    """将异常检测结果导出到 Excel（仅包含被判定为异常的主体）。"""
    if not PANDAS_AVAILABLE:
        print("警告: pandas 未安装，无法导出异常检测 Excel")
        return

    classification = results['classification']
    record_map = {r['project_name']: r for r in results['records']}

    categories = [
        ('两类都异常', 'both_abnormal'),
        ('仅随机森林异常', 'rf_only_abnormal'),
        ('仅硬阈值异常', 'threshold_only_abnormal'),
    ]

    rows = []
    for cat_name, cat_key in categories:
        for unit in classification.get(cat_key, []):
            record = record_map.get(unit, {})
            rows.append({
                '项目名称': unit,
                '所属发电集团': record.get('project_owner', ''),
                '异常类别': cat_name,
                '报价偏离度': round(record.get('price_deviation', 0), 4),
                '收益偏离度': round(record.get('revenue_deviation', 0), 4),
            })

    if rows:
        os.makedirs(os.path.dirname(output_file), exist_ok=True)
        df = pd.DataFrame(rows)
        df.to_excel(output_file, index=False)
        print(f"异常主体已导出到: {output_file}")


# ================================================================
# 市场主体行为聚类（层次聚类）
# ================================================================

# 行为类型名称
ARCHETYPE_COST_ORIENTED = '成本导向型'
ARCHETYPE_AGGRESSIVE_LOW = '激进低价型'
ARCHETYPE_PRICE_ALLIANCE = '价格联盟型'
ARCHETYPE_STRATEGIC_PROBE = '策略试探型'

# 行为类型判定阈值（作用于聚类中心向量）
ALLIANCE_CONSISTENCY_THRESHOLD = 0.0    # 一致性指数 > 0：存在报价影子伙伴，判为价格联盟
AGGRESSIVE_PRICE_DEV_THRESHOLD = -0.20  # 报价偏离度 < -20%：持续大幅低于成本报价
PROBE_CHANGE_RATE_THRESHOLD = 0.30      # |报价变动率| > 30%：跨期大幅调价试探


def classify_cluster_archetype(profile: Dict[str, float]) -> str:
    """
    根据聚类中心向量的阈值规则判定行为类型。
    规则按业务危害优先级排列：联盟 > 低价 > 试探 > 兜底。
    允许多个簇判定为同一类型。
    """
    if profile['bid_consistency_index'] > ALLIANCE_CONSISTENCY_THRESHOLD:
        return ARCHETYPE_PRICE_ALLIANCE
    if profile['price_deviation'] < AGGRESSIVE_PRICE_DEV_THRESHOLD:
        return ARCHETYPE_AGGRESSIVE_LOW
    if profile['abs_bid_change_rate'] > PROBE_CHANGE_RATE_THRESHOLD:
        return ARCHETYPE_STRATEGIC_PROBE
    return ARCHETYPE_COST_ORIENTED


def _adaptive_tree_cut(X: List[List[float]], min_clusters: int = 3, max_clusters: int = 8) -> float:
    """
    自适应确定树状图切分位置（间隙法）。
    构建完整ward树状图，在允许簇数范围内寻找相邻合并高度的最大间隙，
    返回间隙中点作为距离阈值。
    """
    from sklearn.cluster import AgglomerativeClustering

    n = len(X)
    model = AgglomerativeClustering(
        n_clusters=None, distance_threshold=0,
        linkage='ward', compute_full_tree=True
    )
    model.fit(X)
    heights = model.distances_  # 合并高度，升序，长度 n-1

    # 切成 k 类 = 在 heights[n-k-1] 与 heights[n-k] 之间切开
    best_k, best_gap = min_clusters, -1.0
    for k in range(min_clusters, min(max_clusters, n - 1) + 1):
        idx = n - k - 1
        gap = heights[idx + 1] - heights[idx]
        if gap > best_gap:
            best_gap, best_k = gap, k

    idx = n - best_k - 1
    return (heights[idx] + heights[idx + 1]) / 2.0


def hierarchical_clustering(
    records: List[Dict[str, Any]],
    min_clusters: int = 3,
    max_clusters: int = 8
) -> Dict[str, Any]:
    """
    使用层次聚类（AgglomerativeClustering, ward连接）对市场主体行为聚类。

    特征：报价位置系数、报价变动率、报价一致性指数、报价偏离度、
          收益偏离度、报价分布形态特征向量（展开），
          聚类前做z-score标准化消除量纲。
    簇数不预设，采用间隙法在树状图最大合并高度间隙处自适应切分
    （限制在 [min_clusters, max_clusters] 范围内）。

    Args:
        records: 数据记录列表
        min_clusters: 最少簇数（默认2）
        max_clusters: 最多簇数（默认8）

    Returns:
        包含聚类结果的字典
    """
    if not SKLEARN_AVAILABLE:
        return {'error': 'scikit-learn 未安装，无法进行聚类'}

    from sklearn.cluster import AgglomerativeClustering
    from sklearn.preprocessing import StandardScaler

    if len(records) < 2:
        return {'error': f'数据量（{len(records)}）过少，无法聚类'}

    # 复用现有特征提取（PCA压缩bid_curve_vector到1维）
    features, unit_names, feature_names = extract_features(
        records, compress_vector=True, n_components=1
    )

    # 打印PCA信息
    _, explained = compress_bid_curve_pca(records, n_components=1)
    print(f"PCA压缩: bid_curve_vector 6维 → 1维，第1主成分解释方差比例: {explained*100:.1f}%")

    # 标准化（各指标量纲差异大，ward连接基于欧氏距离，必须标准化）
    X = StandardScaler().fit_transform(features)

    # 自适应确定切分位置并聚类
    threshold = _adaptive_tree_cut(X, min_clusters, max_clusters)
    model = AgglomerativeClustering(
        n_clusters=None, distance_threshold=threshold,
        linkage='ward', compute_full_tree=True
    )
    labels = model.fit_predict(X)
    n_clusters = len(set(labels))

    # 按聚类分组
    clusters: Dict[int, List[str]] = {}
    for name, label in zip(unit_names, labels):
        clusters.setdefault(int(label), []).append(name)

    # 计算各聚类中心的关键指标均值（用于行为类型命名）
    record_map = {r['project_name']: r for r in records}
    profiles: Dict[int, Dict[str, float]] = {}
    for label, members in clusters.items():
        n = len(members)
        profiles[label] = {
            'bid_position_coeff': sum(record_map[m]['bid_position_coeff'] for m in members) / n,
            'abs_bid_change_rate': sum(abs(record_map[m]['bid_change_rate']) for m in members) / n,
            'bid_consistency_index': sum(record_map[m]['bid_consistency_index'] for m in members) / n,
            'price_deviation': sum(record_map[m]['price_deviation'] for m in members) / n,
            'revenue_deviation': sum(record_map[m]['revenue_deviation'] for m in members) / n,
        }

    # 对每个簇的中心向量独立做阈值判定，允许多个簇判为同一类型
    archetype_map: Dict[int, str] = {
        label: classify_cluster_archetype(profiles[label])
        for label in clusters
    }

    # 同类型簇编号：价格联盟型-1、价格联盟型-2 ...（仅一个时不编号）
    display_name_map: Dict[int, str] = {}
    type_counters: Dict[str, int] = {}
    type_totals: Dict[str, int] = {}
    for name in archetype_map.values():
        type_totals[name] = type_totals.get(name, 0) + 1
    for label in sorted(clusters.keys()):
        name = archetype_map[label]
        type_counters[name] = type_counters.get(name, 0) + 1
        if type_totals[name] > 1:
            display_name_map[label] = f'{name}-{type_counters[name]}'
        else:
            display_name_map[label] = name

    return {
        'n_clusters': n_clusters,
        'distance_threshold': threshold,
        'labels': dict(zip(unit_names, labels)),
        'clusters': clusters,
        'profiles': profiles,
        'archetype_map': archetype_map,
        'display_name_map': display_name_map,
        'feature_names': feature_names,
    }


def print_clustering_results(cluster_results: Dict[str, Any]):
    """打印聚类结果到控制台。"""
    if 'error' in cluster_results:
        print(f"错误: {cluster_results['error']}")
        return

    print("\n" + "=" * 80)
    print("市场主体行为聚类结果（层次聚类）")
    print("=" * 80)

    clusters = cluster_results['clusters']
    profiles = cluster_results['profiles']
    display_name_map = cluster_results['display_name_map']

    print(f"\n使用特征: {', '.join(cluster_results['feature_names'])}")
    print(f"自适应簇数: {cluster_results['n_clusters']}（切分距离阈值: "
          f"{cluster_results['distance_threshold']:.4f}），总主体数: "
          f"{sum(len(m) for m in clusters.values())}")
    print(f"判定阈值: 一致性指数>{ALLIANCE_CONSISTENCY_THRESHOLD} → 价格联盟型 | "
          f"报价偏离度<{AGGRESSIVE_PRICE_DEV_THRESHOLD} → 激进低价型 | "
          f"|报价变动率|>{PROBE_CHANGE_RATE_THRESHOLD} → 策略试探型 | "
          f"其余 → 成本导向型")

    # 按行为类型优先级排序输出
    order = [ARCHETYPE_PRICE_ALLIANCE, ARCHETYPE_AGGRESSIVE_LOW,
             ARCHETYPE_STRATEGIC_PROBE, ARCHETYPE_COST_ORIENTED]
    archetype_map = cluster_results['archetype_map']
    sorted_labels = sorted(clusters.keys(),
                           key=lambda l: (order.index(archetype_map[l])
                                          if archetype_map[l] in order else 99, l))

    for label in sorted_labels:
        members = clusters[label]
        profile = profiles[label]
        archetype = display_name_map[label]
        print(f"\n【{archetype}】（簇 {label}，{len(members)} 个主体）")
        print(f"  聚类中心特征:")
        print(f"    报价位置系数: {profile['bid_position_coeff']:.4f}")
        print(f"    报价变动率(绝对值): {profile['abs_bid_change_rate']:.4f}")
        print(f"    报价一致性指数: {profile['bid_consistency_index']:.4f}")
        print(f"    报价偏离度: {profile['price_deviation']:.4f} ({profile['price_deviation']*100:.2f}%)")
        print(f"    收益偏离度: {profile['revenue_deviation']:.4f} ({profile['revenue_deviation']*100:.2f}%)")
        print(f"  成员:")
        for m in sorted(members):
            print(f"    - {m}")

    print("\n" + "=" * 80)


def prompt_cluster_selection(cluster_results: Dict[str, Any]) -> Optional[List[int]]:
    """
    人机交互：让用户选择要进行异常行为特征识别的簇。

    Returns:
        选中的簇号列表，用户退出时返回None
    """
    clusters = cluster_results['clusters']
    display_name_map = cluster_results['display_name_map']

    print("\n可选择的簇：")
    for label in sorted(clusters.keys()):
        print(f"  簇 {label}: {display_name_map[label]}（{len(clusters[label])} 个主体）")

    while True:
        choice = input("\n请输入要进行异常行为特征识别的簇号"
                       "（多个用逗号分隔，输入 q 退出）: ").strip()
        if choice.lower() == 'q':
            return None
        try:
            selected = [int(x.strip()) for x in choice.replace('，', ',').split(',') if x.strip()]
        except ValueError:
            print("输入格式错误，请输入数字簇号，多个用逗号分隔")
            continue
        invalid = [c for c in selected if c not in clusters]
        if not selected:
            print("未输入簇号，请重新输入")
            continue
        if invalid:
            print(f"簇号 {invalid} 不存在，请重新输入")
            continue
        return selected


if __name__ == "__main__":
    import sys

    # 默认数据文件路径
    root_dir = os.path.dirname(os.path.dirname(__file__))
    default_data_file = os.path.join(
        root_dir, 'Data', 'unit_behavior_analysis', 'bidding_analysis_results.csv'
    )

    # 支持命令行参数指定文件
    if len(sys.argv) > 1:
        data_file = sys.argv[1]
    else:
        data_file = default_data_file

    if not os.path.exists(data_file):
        print(f"错误: 数据文件不存在: {data_file}")
        sys.exit(1)

    print(f"加载数据文件: {data_file}")

    # 步骤1：加载数据并进行市场主体分类（层次聚类）
    records = load_analysis_data(data_file)
    if not records:
        print("错误: 未找到数据或文件为空")
        sys.exit(1)
    print(f"加载了 {len(records)} 条记录")

    cluster_results = hierarchical_clustering(records)
    print_clustering_results(cluster_results)
    if 'error' in cluster_results:
        sys.exit(1)

    # 步骤2：人机交互选择要进行异常识别的簇
    selected_labels = prompt_cluster_selection(cluster_results)
    if selected_labels is None:
        print("已退出")
        sys.exit(0)

    # 步骤3：用选中簇的市场主体作为异常行为特征识别的输入
    clusters = cluster_results['clusters']
    display_name_map = cluster_results['display_name_map']
    selected_names = set()
    for label in selected_labels:
        selected_names.update(clusters[label])
    selected_records = [r for r in records if r['project_name'] in selected_names]

    print(f"\n已选择簇 {selected_labels}（"
          f"{'、'.join(display_name_map[l] for l in selected_labels)}），"
          f"共 {len(selected_records)} 个市场主体")

    # 步骤4：异常行为特征识别并打印结果
    results = run_abnormal_detection(selected_records)
    print_detection_results(results)
