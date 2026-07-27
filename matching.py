"""
备件/物料匹配核心算法 - 基于注意力机制的模糊匹配
实现从匹配算法核心逻辑文档提取的 attention-based matching
"""

import re
import math
from collections import Counter

try:
    import jieba
except ImportError:
    jieba = None


def clean_str(s) -> str:
    """字符串清洗：标准化输入，消除格式差异"""
    if s is None or (isinstance(s, float) and math.isnan(s)):
        return ""
    s = str(s).upper()  # 转大写
    # 去除全角空格、零宽字符、制表符、换行等
    s = re.sub(r"[　\xa0​‌‍\t\n\r]", " ", s)
    # 统一分隔符为空格
    s = re.sub(r"[_\-／/．.。,，]", " ", s)
    # 合并多个空格
    s = re.sub(r"\s+", " ", s).strip()
    return s


def build_ngram_index(candidates: list, n: int = 3) -> dict:
    """
    构建 n-gram 倒排索引
    对每个候选串提取所有长度为 n 的子串，建立倒排映射
    """
    index = {}
    for idx, text in enumerate(candidates):
        seen = set()
        for i in range(len(text) - n + 1):
            gram = text[i : i + n]
            if gram not in seen:
                seen.add(gram)
                if gram not in index:
                    index[gram] = []
                index[gram].append(idx)
    return index


def get_candidate_indices(query: str, index: dict, n: int) -> list:
    """
    通过 n-gram 预过滤，筛选出候选集
    """
    query_grams = set()
    for i in range(len(query) - n + 1):
        query_grams.add(query[i : i + n])

    if not query_grams:
        return []

    # 统计每个候选的命中 n-gram 数
    hit_counts = Counter()
    for gram in query_grams:
        if gram in index:
            for idx in index[gram]:
                hit_counts[idx] += 1

    if not hit_counts:
        return []

    max_hits = max(hit_counts.values())
    # 取命中数 >= max_hits * 0.3 的候选
    threshold = max(1, max_hits * 0.3)
    return [idx for idx, cnt in hit_counts.items() if cnt >= threshold]


def build_idf_dict(all_texts: list) -> dict:
    """
    构建 IDF 词典 - 计算每个词的稀有度
    IDF(t) = ln((N+1)/(DF(t)+1)) + 1
    """
    N = len(all_texts)
    df = {}
    for text in all_texts:
        seen = set()
        # jieba 分词
        if jieba:
            for token in jieba.cut(text):
                token = token.strip()
                if token and token not in seen:
                    seen.add(token)
                    df[token] = df.get(token, 0) + 1
        # 字符 2-gram
        for i in range(len(text) - 1):
            token = text[i : i + 2]
            if token.strip() and token not in seen:
                seen.add(token)
                df[token] = df.get(token, 0) + 1

    idf = {}
    for token, doc_freq in df.items():
        idf[token] = math.log((N + 1) / (doc_freq + 1)) + 1
    return idf


def attention_score(query: str, candidate: str, idf_dict: dict,
                    model_text: str = "") -> float:
    """
    注意力评分核心函数
    返回 0-100 的匹配度分数
    model_text: 规格型号原文，单独用于型号匹配奖励（最准确列，权重更高）
    """
    if not query or not candidate:
        return 0.0

    # --- Token 提取 ---
    tokens = []
    # jieba 中文分词
    if jieba:
        for token in jieba.cut(query):
            token = token.strip()
            if token:
                tokens.append(token)
    else:
        tokens.append(query)

    # 字符 2-gram
    for i in range(len(query) - 1):
        gram = query[i : i + 2]
        if gram.strip():
            tokens.append(gram)

    if not tokens:
        return 0.0

    # --- 加权评分 ---
    total_weight = 0.0
    total_score = 0.0

    for pos, token in enumerate(tokens):
        # Token 权重 = IDF × 位置衰减
        token_weight = idf_dict.get(token, 1.0) * (0.95**pos)

        # Token 得分
        if token in candidate:
            # 完全匹配
            tf = candidate.count(token)
            tf_factor = 1.0 + 0.1 * min(tf - 1, 3)  # 最多 +0.3
            token_score = 1.0 * tf_factor
        elif len(token) >= 2:
            # 部分匹配：字符重叠率
            overlap = sum(1 for c in token if c in candidate)
            token_score = (overlap / len(token)) * 0.8
        else:
            token_score = 0.0

        total_weight += token_weight
        total_score += token_score * token_weight

    if total_weight == 0:
        return 0.0

    base_score = (total_score / total_weight) * 100

    # --- 完整包含奖励 ---
    bonus = 0.0
    if query in candidate:
        bonus = 5.0 * (1 - len(query) / max(len(candidate), 1))
    elif candidate in query:
        bonus = 3.0

    # --- 长度差异惩罚 ---
    len_penalty = 0.0
    if len(query) < len(candidate) * 0.3:
        len_penalty = 5.0 * (1 - len(query) / max(len(candidate), 1))

    # --- 型号特异性扣分 ---
    spec_penalty = 0.0
    model_tokens = re.findall(r"[A-Z0-9]{6,}", query)
    if model_tokens:
        matched_model_count = sum(
            1 for mt in model_tokens if any(mt[:4] in candidate or mt in candidate for mt in [mt])
        )
        missing = len(model_tokens) - matched_model_count
        if missing > 0 and matched_model_count == 0:
            spec_penalty = 15.0 * min(missing, 3)

    # --- 型号匹配奖励（型号列是最准确的匹配信号）---
    model_bonus = 0.0
    if model_text:
        mt_clean = clean_str(str(model_text))
        if mt_clean.strip():
            mt_tokens = set()
            if jieba:
                for token in jieba.cut(mt_clean):
                    token = token.strip()
                    if token:
                        mt_tokens.add(token)
            for i in range(len(mt_clean) - 1):
                gram = mt_clean[i:i+2]
                if gram.strip():
                    mt_tokens.add(gram)
            if mt_tokens:
                match_count = sum(1 for t in mt_tokens if t in candidate)
                ratio = match_count / len(mt_tokens)
                model_bonus = ratio * 15.0  # 最高 +15 分

    final_score = base_score + bonus - len_penalty - spec_penalty + model_bonus
    return max(0.0, min(100.0, final_score))


def find_best_match(
    query_text: str,
    candidates: list,
    idf_dict: dict,
    index: dict = None,
    n: int = 3,
    threshold: float = 50.0,
) -> tuple:
    """
    在候选中找到最佳匹配
    返回 (最佳索引, 最佳分数, 是否通过阈值)
    """
    # 预处理查询
    query_clean = clean_str(query_text)

    # 确定候选范围
    if index:
        candidate_indices = get_candidate_indices(query_clean, index, n)
        if candidate_indices is None:
            candidate_indices = list(range(len(candidates)))
    else:
        candidate_indices = list(range(len(candidates)))

    best_score = 0.0
    best_idx = -1

    for idx in candidate_indices:
        score = attention_score(query_clean, candidates[idx], idf_dict)
        if score > best_score:
            best_score = score
            best_idx = idx

    return best_idx, best_score, best_score >= threshold


def build_candidate_text(name: str, model: str = "") -> str:
    """
    构建候选文本字符串用于匹配
    仅使用物料名称 + 型号（图片文件名中仅包含这两部分信息）
    """
    parts = [clean_str(name)]
    if model and str(model).strip():
        parts.append(clean_str(str(model)))
    return " ".join(parts)
