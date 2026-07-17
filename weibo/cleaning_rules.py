# -*- coding: utf-8 -*-
import re


INDUSTRY_KEYWORDS = [
    '马克水印相机',
    '水印照片',
    '水印相机',
    '水印打卡',
    '今日水印相机',
    '元道经纬',
    '外勤记录',
    '外勤打卡',
    '外勤相机',
]

STRONG_PRODUCT_TERMS = [
    '马克水印相机',
    '今日水印相机',
    '元道经纬',
    '水印相机',
    '外勤相机',
    '水印打卡',
    '外勤打卡',
    '外勤记录',
]

SCENARIO_TERMS = [
    '外勤',
    '打卡',
    '考勤',
    '签到',
    '定位',
    '时间',
    '地点',
    '地址',
    '经纬度',
    '现场',
    '工地',
    '工程',
    '施工',
    '巡检',
    '巡店',
    '门店',
    '物业',
    '保洁',
    '家政',
    '销售',
    '拜访',
    '客户',
    '团队',
    '管理',
    '记录',
    '拍照',
    '上传',
    '审批',
    '日报',
]

FEEDBACK_TERMS = [
    '问题',
    '反馈',
    '需求',
    '建议',
    '吐槽',
    '投诉',
    '希望',
    '能不能',
    '不能',
    '无法',
    '失败',
    '报错',
    '卡',
    '闪退',
    '崩溃',
    '定位不准',
    '水印不准',
    '时间不对',
    '打不开',
    '不好用',
    '怎么',
    '求',
]

EXCLUSION_RULES = [
    {
        'name': 'fan_media_or_celebrity',
        'reason': '粉圈、明星、超话或饭拍内容，通常只是在声明图片水印/侵删。',
        'patterns': [
            '超话', '饭拍', '直拍', '生图', '站姐', '应援', '控评', '打榜', '安利',
            '爱豆', '偶像', '粉丝', '明星', '演员', 'TFBOYS', '易烊千玺', '王一博',
            '肖战', '时代少年团', '张极', '张泽禹', '迪丽热巴', '杨幂', '白鹿'
        ],
    },
    {
        'name': 'wallpaper_avatar_material',
        'reason': '头像、壁纸、素材、无水印原图类内容，不是水印相机产品反馈。',
        'patterns': [
            '头像', '壁纸', '无水印', '原图', '套图', '修图', '拼图', '美图',
            '素材', '表情包', 'plog', '写真', '高清图', '水印侵删', '侵删'
        ],
    },
    {
        'name': 'general_photo_or_camera',
        'reason': '泛照片/手机相机/摄影内容，缺少外勤、打卡、管理或产品反馈语境。',
        'patterns': [
            '摄影', '拍摄', '相册', '照片墙', '镜头', '手机影像', '人像', '滤镜',
            '构图', '风景照', '旅拍', '随手拍'
        ],
    },
    {
        'name': 'poetry_or_long_text_noise',
        'reason': '长文、诗文、玄学或营销噪音，命中关键词但行业关联弱。',
        'patterns': [
            '三伏天', '忌日', '梅花易数', '天元', '地元', '乾坤', '阴阳',
            '博喻', '排比', '顺其自然'
        ],
    },
]


def _contains_any(text, terms):
    return [term for term in terms if term and term in text]


def _matched_exclusions(text):
    matches = []
    for rule in EXCLUSION_RULES:
        matched = _contains_any(text, rule['patterns'])
        if matched:
            matches.append({
                'name': rule['name'],
                'reason': rule['reason'],
                'matched_terms': matched,
            })
    return matches


def evaluate_weibo(weibo, keyword):
    """Score one Weibo item for watermark-camera industry monitoring."""
    text = ''.join([
        weibo.get('text', '') or '',
        ' ',
        weibo.get('topics', '') or '',
        ' ',
        weibo.get('source', '') or '',
        ' ',
        keyword or '',
    ])
    compact_text = re.sub(r'\s+', '', text)
    strong_terms = _contains_any(compact_text, STRONG_PRODUCT_TERMS)
    scenario_terms = _contains_any(compact_text, SCENARIO_TERMS)
    feedback_terms = _contains_any(compact_text, FEEDBACK_TERMS)
    exclusion_matches = _matched_exclusions(compact_text)

    score = 0
    tags = []
    if strong_terms:
        score += 3
        tags.append('产品/品牌命中')
    if scenario_terms:
        score += 2
        tags.append('外勤/打卡场景')
    if feedback_terms:
        score += 2
        tags.append('客户反馈/需求')
    if keyword == '水印照片' and not (strong_terms or scenario_terms or feedback_terms):
        score -= 2
    if exclusion_matches:
        score -= 3

    # Strong product and field-work context can rescue a post that also contains
    # broad words like photo or wallpaper.
    if strong_terms and (scenario_terms or feedback_terms):
        score += 2
        exclusion_matches = []

    is_related = score >= 2 and not exclusion_matches
    if is_related:
        reason = '保留：' + '、'.join(tags or ['命中行业关键词'])
    elif exclusion_matches:
        reason = '排除：' + '；'.join(match['reason'] for match in exclusion_matches)
    else:
        reason = '排除：只命中宽泛关键词，缺少外勤打卡、产品使用或客户反馈语境。'

    return {
        'is_related': is_related,
        'score': score,
        'tags': tags,
        'reason': reason,
        'matched_product_terms': strong_terms,
        'matched_scenario_terms': scenario_terms,
        'matched_feedback_terms': feedback_terms,
        'exclusion_rules': [match['name'] for match in exclusion_matches],
    }
