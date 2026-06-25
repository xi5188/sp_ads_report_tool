import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime

# ---------------------- 全局映射字典 ----------------------
bid_name_map = {
    "Fixed bid": "固定竞价",
    "Dynamic bids - up and down": "提高和降低",
    "Dynamic bids - down only": "仅降低"
}
placement_name_map = {
    "Placement Amazon Business": "企业购广告位",
    "Placement Rest Of Search": "搜索结果的其余位置",
    "Placement Product Page": "商品页面",
    "Placement Top": "搜索结果顶部（首页）"
}
entity_map = {
    "Campaign": "广告活动",
    "Bidding Adjustment": "广告位调价",
    "Ad Group": "广告组",
    "Product Ad": "商品广告",
    "Product Targeting": "ASIN定位投放",
    "Keyword": "投放关键词",
    "Negative Keyword": "否定关键词",
    "Negative Product Targeting": "否定ASIN",
    "Campaign Negative Keyword": "广告活动否定关键词"
}
status_map = {
    "enabled": "已启用",
    "paused": "已暂停",
    "Eligible": "符合条件",
    "Data not available": "数据不可用",
    "Ineligible for ad creation": "不符合条件"
}
match_type_map = {
    "Broad": "广泛",
    "Phrase": "词组",
    "Exact": "精准",
    "Negative Phrase": "否定词组",
    "Negative Exact": "否定精准",
    "None": "无"
}
# 新增：定投类目表达式中英映射
target_expr_map = {
    "close-match": "紧密匹配",
    "loose-match": "宽泛匹配",
    "complements": "关联商品",
    "substitutes": "同类商品"
}

# 文件读取缓存函数（大文件优化核心）
@st.cache_data(max_entries=5, ttl=3600)
def read_sp_excel(file_obj):
    sheets = pd.read_excel(file_obj, sheet_name=None)
    df_ad = None
    df_search = None

    # ========== 广告表：仅加载代码用到的字段，剔除无用列 ==========
    if "Sponsored Products Campaigns" in sheets:
        raw_ad = sheets["Sponsored Products Campaigns"]
        raw_ad.columns = raw_ad.columns.str.strip()
        # 程序全部用到的列，多余列直接丢弃
        keep_ad_cols = [
            "Portfolio Name (Informational only)",
            "Campaign Name (Informational only)",
            "Ad Group Name (Informational only)",
            "Entity",
            "State",
            "Daily Budget",
            "Bidding Strategy",
            "Placement",
            "Percentage",
            "Bid",
            "Product Targeting Expression",
            "Keyword Text",
            "Match Type",
            "Eligibility Status (Informational only)",
            "SKU",
            "ASIN (Informational only)",
            "Impressions",
            "Clicks",
            "Spend",
            "Sales",
            "Orders",
            "Units",
            "Click-through Rate",
            "Conversion Rate",
            "ACOS",
            "CPC",
            "ROAS"
        ]
        # 过滤掉报表不存在的列，防止报错
        use_cols_ad = [col for col in keep_ad_cols if col in raw_ad.columns]
        raw_ad = raw_ad[use_cols_ad].copy()
        # 数值统一转换
        num_cols = ["Impressions","Clicks","Spend","Sales","Orders","Units","Percentage","Bid","Click-through Rate","Conversion Rate","ACOS","CPC","ROAS"]
        for c in num_cols:
            if c in raw_ad.columns:
                raw_ad[c] = pd.to_numeric(raw_ad[c], errors="coerce").fillna(0)
        raw_ad["Ad Group Name (Informational only)"] = raw_ad["Ad Group Name (Informational only)"].fillna("无广告组")
        df_ad = raw_ad

    # ========== 搜索词表：精简字段 ==========
    if "SP Search Term Report" in sheets:
        raw_st = sheets["SP Search Term Report"]
        raw_st.columns = raw_st.columns.str.strip()
        keep_st_cols = [
            "Campaign Name (Informational only)",
            "Campaign Name",
            "Keyword Text",
            "Match Type",
            "Product Targeting Expression",
            "Customer Search Term",
            "Search Term",
            "Impressions",
            "Clicks",
            "Spend",
            "Sales",
            "Orders",
            "Units",
            "Click-through Rate",
            "Conversion Rate",
            "ACOS",
            "CPC",
            "ROAS"
        ]
        use_cols_st = [col for col in keep_st_cols if col in raw_st.columns]
        raw_st = raw_st[use_cols_st].copy()
        search_num = ["Impressions","Clicks","Spend","Sales","Orders","Units","Click-through Rate","Conversion Rate","ACOS","CPC","ROAS"]
        for c in search_num:
            if c in raw_st.columns:
                raw_st[c] = pd.to_numeric(raw_st[c], errors="coerce").fillna(0)
        df_search = raw_st
    return df_ad, df_search

# 页面基础配置
st.set_page_config(page_title="亚马逊SP广告数据看板", layout="wide", page_icon="📊")
st.markdown("""
<style>
    .metric-box {background:#f0f7ff;padding:16px;border-radius:12px;border-left:5px #1677ff solid;text-align:center;}
    .block-title {font-size:14px;font-weight:bold;background:#f5f7f9;padding:8px;border-radius:4px;margin-top:12px;}
    .search-bottom-panel {margin-top:20px;border-top:2px solid #ddd;padding-top:20px;}
</style>
""", unsafe_allow_html=True)

# 会话缓存初始化（新增存储关键词三重匹配字段 + 定投表达式缓存）
if "df_ad" not in st.session_state:
    st.session_state.df_ad = None
if "df_search" not in st.session_state:
    st.session_state.df_search = None
if "sel_campaign_name" not in st.session_state:
    st.session_state.sel_campaign_name = ""
# 关键词三重匹配缓存
if "sel_keyword_text" not in st.session_state:
    st.session_state.sel_keyword_text = ""
if "sel_match_type_raw" not in st.session_state:
    st.session_state.sel_match_type_raw = ""
# ASIN定投匹配缓存
if "sel_target_expr" not in st.session_state:
    st.session_state.sel_target_expr = ""

# ===================== 页面顶部：主标题 + 文件上传 =====================
st.title("📊 亚马逊SP广告看板｜原始明细匹配工具")
st.divider()

upload_xlsx = st.file_uploader("上传Excel文件（Sponsored Products Campaigns + SP Search Term Report）", type=["xlsx"])
if upload_xlsx:
    try:
        # 调用缓存函数解析文件，重复上传相同文件不会重复解析
        df_raw_ad, df_raw_search = read_sp_excel(upload_xlsx)

        if df_raw_ad is not None:
            st.session_state.df_ad = df_raw_ad
            st.success("✅ 广告明细加载完成")
        else:
            st.error("缺少工作表：Sponsored Products Campaigns")

        st.session_state.df_search = df_raw_search
        if df_raw_search is not None:
            st.success("✅ 搜索词报表加载完成")
        else:
            st.warning("未上传SP Search Term Report，搜索词模块隐藏")
        st.divider()
    except Exception as e:
        st.error(f"文件读取失败：{str(e)}")

df_ad = st.session_state.df_ad
df_search = st.session_state.df_search

# 全局统一百分比格式化函数
def format_percent(x):
    if pd.isna(x) or x == 0:
        return "0.00%"
    return f"{x:.2%}"

if df_ad is not None:
    # ===================== 一、全账户投放总览（顶部，仅统计Entity=Campaign，两行3卡片） =====================
    st.subheader("一、全账户投放总览（全局合计）")
    filter_df = df_ad.copy()
    camp_only_data = filter_df[filter_df["Entity"] == "Campaign"].copy()
    sum_imp = camp_only_data["Impressions"].sum()
    sum_click = camp_only_data["Clicks"].sum()
    sum_spend = camp_only_data["Spend"].sum()
    sum_sales = camp_only_data["Sales"].sum()
    sum_order = camp_only_data["Orders"].sum()

    # 比率指标重新计算
    ctr = sum_click / sum_imp if sum_imp > 0 else 0
    cvr = sum_order / sum_click if sum_click > 0 else 0
    acos = sum_spend / sum_sales if sum_sales > 0 else 0
    roas = sum_sales / sum_spend if sum_spend > 0 else 0

    row1 = st.columns(3)
    with row1[0]:
        st.markdown(f"<div class='metric-box'><div>总曝光</div><div style='font-size:24px;font-weight:bold'>{int(sum_imp):,}</div></div>", unsafe_allow_html=True)
    with row1[1]:
        st.markdown(f"<div class='metric-box'><div>总点击</div><div style='font-size:24px;font-weight:bold'>{int(sum_click):,}</div></div>", unsafe_allow_html=True)
    with row1[2]:
        st.markdown(f"<div class='metric-box'><div>总花费</div><div style='font-size:24px;font-weight:bold'>${sum_spend:,.2f}</div></div>", unsafe_allow_html=True)
    row2 = st.columns(3)
    with row2[0]:
        st.markdown(f"<div class='metric-box'><div>总销售额</div><div style='font-size:24px;font-weight:bold'>${sum_sales:,.2f}</div></div>", unsafe_allow_html=True)
    with row2[1]:
        st.markdown(f"<div class='metric-box'><div>整体ACOS</div><div style='font-size:24px;font-weight:bold'>{format_percent(acos)}</div></div>", unsafe_allow_html=True)
    with row2[2]:
        st.markdown(f"<div class='metric-box'><div>ROAS</div><div style='font-size:24px;font-weight:bold'>{roas:.2f}</div></div>", unsafe_allow_html=True)
    st.divider()

    # ===================== 二、广告筛选（三控件同一行） =====================
    st.subheader("二、广告筛选")
    col_name, col_entity, col_state = st.columns([4, 2, 2])
    # 实体类型下拉：全部 / 广告活动 / 广告组合
    with col_entity:
        entity_sel_cn = st.selectbox("实体类型", ["全部", "广告活动", "广告组合"])
    with col_name:
        # 根据实体下拉切换检索提示
        if entity_sel_cn == "广告活动":
            filter_text = st.text_input("名称检索（广告活动）")
        elif entity_sel_cn == "广告组合":
            filter_text = st.text_input("名称检索（广告组合）")
        else:
            filter_text = st.text_input("名称检索（全部实体通用）")
    # 投放状态下拉：中文展示，映射后台英文
    with col_state:
        state_cn_list = ["全部", "已启用", "已暂停"]
        state_sel_cn = st.selectbox("投放状态", state_cn_list)
        # 映射后台真实字段值
        if state_sel_cn == "已启用":
            state_sel = "enabled"
        elif state_sel_cn == "已暂停":
            state_sel = "paused"
        else:
            state_sel = "全部"

    # 执行筛选逻辑
    filter_df = df_ad.copy()
    # 实体过滤
    if entity_sel_cn == "广告活动":
        filter_df = filter_df[filter_df["Entity"] == "Campaign"]
    elif entity_sel_cn == "广告组合":
        pass
    # 名称检索过滤
    if filter_text.strip() != "":
        if entity_sel_cn == "广告活动":
            filter_df = filter_df[filter_df["Campaign Name (Informational only)"].str.contains(filter_text, na=False, case=False)]
        elif entity_sel_cn == "广告组合":
            filter_df = filter_df[filter_df["Portfolio Name (Informational only)"].astype(str).str.contains(filter_text, na=False, case=False)]
    # 投放状态过滤
    if state_sel != "全部":
        filter_df = filter_df[filter_df["State"] == state_sel]
    st.divider()

    # ===================== 三、广告活动总表（默认Spend花费降序） =====================
    campaign_df = filter_df[filter_df["Entity"] == "Campaign"].copy()
    campaign_df = campaign_df.sort_values("Spend", ascending=False, ignore_index=True)
    disp_raw = campaign_df.copy()
    disp_raw["Daily Budget"] = disp_raw["Daily Budget"].apply(lambda x: f"${x:.2f}" if pd.notna(x) else "$0.00")
    disp_raw["Spend"] = disp_raw["Spend"].apply(lambda x: f"${x:.2f}")
    disp_raw["Sales"] = disp_raw["Sales"].apply(lambda x: f"${x:.2f}")
    disp_raw["CPC"] = disp_raw["CPC"].apply(lambda x: f"${x:.2f}" if pd.notna(x) else "$0.00")
    disp_raw["Click-through Rate"] = disp_raw["Click-through Rate"].apply(format_percent)
    disp_raw["Conversion Rate"] = disp_raw["Conversion Rate"].apply(format_percent)
    disp_raw["ACOS"] = disp_raw["ACOS"].apply(format_percent)
    disp_raw["ROAS"] = disp_raw["ROAS"].apply(lambda x: f"{x:.2f}" if pd.notna(x) else "0.00")
    disp_raw["Bidding Strategy"] = disp_raw["Bidding Strategy"].map(bid_name_map).fillna(disp_raw["Bidding Strategy"])
    disp_raw["State"] = disp_raw["State"].map(status_map).fillna(disp_raw["State"])

    camp_show_cols = [
        "Portfolio Name (Informational only)",
        "Campaign Name (Informational only)","State","Daily Budget","Bidding Strategy",
        "Impressions","Clicks","Click-through Rate","Spend","Sales","Orders","Units","Conversion Rate","ACOS","CPC","ROAS"
    ]
    camp_display = disp_raw[camp_show_cols].copy()
    camp_display.columns = [
        "广告组合", "广告活动名称","投放状态","日预算($)","竞价策略",
        "总曝光","总点击","点击率","花费($)","销售额($)","订单量","Units","转化率","ACOS","CPC","ROAS"
    ]
    st.subheader("三、广告活动总表（单击单行查看下方分层明细）")
    cfg_campaign = {col: st.column_config.Column(width="stretch") for col in camp_display.columns}
    camp_select = st.dataframe(
        camp_display,
        use_container_width=True,
        height=390,
        on_select="rerun",
        selection_mode="single-row",
        column_config=cfg_campaign
    )
    selected_rows = camp_select.selection.rows
    if len(selected_rows) > 0:
        sel_idx = selected_rows[0]
        st.session_state.sel_campaign_name = camp_display.iloc[sel_idx]["广告活动名称"]
        # 切换广告时清空关键词+定投全部匹配缓存
        st.session_state.sel_keyword_text = ""
        st.session_state.sel_match_type_raw = ""
        st.session_state.sel_target_expr = ""
    else:
        st.session_state.sel_campaign_name = ""
        st.session_state.sel_keyword_text = ""
        st.session_state.sel_match_type_raw = ""
        st.session_state.sel_target_expr = ""
    st.divider()

    # ===================== 四、当前广告分层明细 =====================
    st.subheader("四、当前广告分层明细")
    sel_camp_name = st.session_state.sel_campaign_name
    if sel_camp_name == "":
        st.info("请单击上方广告活动表格任意一行选中广告")
    else:
        st.markdown(f"<p style='color:#1677ff'>已选中广告：{sel_camp_name}</p>", unsafe_allow_html=True)
        current_camp_data = df_ad[df_ad["Campaign Name (Informational only)"] == sel_camp_name]

        # 1.广告位调价模块 height=180
        with st.container():
            bid_data = current_camp_data[current_camp_data["Entity"] == "Bidding Adjustment"]
            camp_bid_raw = current_camp_data["Bidding Strategy"].drop_duplicates().dropna()
            bid_cn = bid_name_map[camp_bid_raw.iloc[0]] if len(camp_bid_raw) > 0 else "无"
            if len(bid_data) > 0:
                st.markdown("<div class='block-title'>1. 广告位调价 Bidding Adjustment</div>", unsafe_allow_html=True)
                bid_cols = ["Bidding Strategy","Placement","Percentage","Impressions","Clicks","Click-through Rate","Spend","Sales","Orders","Units","Conversion Rate","ACOS","CPC","ROAS"]
                bid_view = bid_data[bid_cols].copy()
                bid_view["Bidding Strategy"] = bid_cn
                bid_view["Placement"] = bid_view["Placement"].map(placement_name_map).fillna(bid_view["Placement"])
                bid_view["Percentage"] = bid_view["Percentage"].apply(lambda x: f"{float(x):.0f}%" if pd.notna(x) else "0%")
                bid_view["Click-through Rate"] = bid_view["Click-through Rate"].apply(format_percent)
                bid_view["Conversion Rate"] = bid_view["Conversion Rate"].apply(format_percent)
                bid_view["ACOS"] = bid_view["ACOS"].apply(format_percent)
                bid_view.columns = [
                    "竞价策略","广告位","百分比","曝光","点击","点击率","花费($)","销售额($)","订单量","销量","转化率","ACOS","CPC","ROAS"
                ]
                cfg_bid = {c: st.column_config.Column(width="stretch") for c in bid_view.columns}
                st.dataframe(bid_view, use_container_width=True, height=180, column_config=cfg_bid)
                st.divider()

        # 2.广告组汇总总表
        with st.container():
            st.markdown("<div class='block-title'>2. 广告组 Ad Group 汇总</div>", unsafe_allow_html=True)
            adgroup_df = current_camp_data[current_camp_data["Entity"] == "Ad Group"].copy()
            if len(adgroup_df) > 0:
                ag_show_cols = [
                    "Ad Group Name (Informational only)","State","Impressions","Clicks","Click-through Rate",
                    "Spend","Sales","Orders","Units","Conversion Rate","ACOS","CPC","ROAS"
                ]
                ag_view = adgroup_df[ag_show_cols].copy()
                ag_view["Spend"] = ag_view["Spend"].apply(lambda x: f"${x:.2f}")
                ag_view["Sales"] = ag_view["Sales"].apply(lambda x: f"${x:.2f}")
                ag_view["CPC"] = ag_view["CPC"].apply(lambda x: f"${x:.2f}" if pd.notna(x) else "$0.00")
                ag_view["Click-through Rate"] = ag_view["Click-through Rate"].apply(format_percent)
                ag_view["Conversion Rate"] = ag_view["Conversion Rate"].apply(format_percent)
                ag_view["ACOS"] = ag_view["ACOS"].apply(format_percent)
                ag_view["State"] = ag_view["State"].map(status_map).fillna(ag_view["State"])
                ag_view.columns = [
                    "广告组名称","投放状态","曝光","点击","点击率",
                    "花费($)","销售额($)","订单量","Units","转化率","ACOS","CPC","ROAS"
                ]
                cfg_ag_total = {c: st.column_config.Column(width="stretch") for c in ag_view.columns}
                st.dataframe(ag_view, use_container_width=True, height=110, column_config=cfg_ag_total)
            else:
                st.info("该广告下无广告组汇总数据")
            st.divider()

        # 3.广告组细分折叠面板
        ag_name_list = current_camp_data["Ad Group Name (Informational only)"].drop_duplicates().unique()
        ag_name_list = [name for name in ag_name_list if name != "无广告组"]
        if len(ag_name_list) > 0:
            for ag_name in ag_name_list:
                ag_data = current_camp_data[current_camp_data["Ad Group Name (Informational only)"] == ag_name]
                with st.expander(f"广告组：{ag_name}", expanded=False):
                    # 商品广告
                    with st.container():
                        prod = ag_data[ag_data["Entity"] == "Product Ad"]
                        if len(prod) > 0:
                            st.markdown("<div class='block-title'>▸ 商品广告 Product Ad</div>", unsafe_allow_html=True)
                            p_cols = ["Eligibility Status (Informational only)","SKU","ASIN (Informational only)","Impressions","Clicks","Click-through Rate","Spend","Sales","Orders","Units","Conversion Rate","ACOS","CPC","ROAS"]
                            p_view = prod[p_cols].copy()
                            p_view["Eligibility Status (Informational only)"] = p_view["Eligibility Status (Informational only)"].map(status_map).fillna(p_view["Eligibility Status (Informational only)"])
                            p_view["Click-through Rate"] = p_view["Click-through Rate"].apply(format_percent)
                            p_view["Conversion Rate"] = p_view["Conversion Rate"].apply(format_percent)
                            p_view["ACOS"] = p_view["ACOS"].apply(format_percent)
                            p_view.columns = [
                                "投放状态","SKU","ASIN","曝光","点击","点击率","花费($)","销售额($)","订单量","销量","转化率","ACOS","CPC","ROAS"
                            ]
                            cfg_prod = {c: st.column_config.Column(width="stretch") for c in p_view.columns}
                            st.dataframe(p_view, use_container_width=True, height=110, column_config=cfg_prod)
                    # ASIN定位投放（新增：表格翻译定投类目，缓存取值从原始pt拿英文原值）
                    with st.container():
                        pt = ag_data[ag_data["Entity"] == "Product Targeting"]
                        if len(pt) > 0:
                            st.markdown("<div class='block-title'>▸ ASIN定位投放（单击行筛选下方搜索词）</div>", unsafe_allow_html=True)
                            pt_cols = ["State","Product Targeting Expression","Bid","Impressions","Clicks","Click-through Rate","Spend","Sales","Orders","Units","Conversion Rate","ACOS","CPC","ROAS"]
                            pt_view = pt[pt_cols].copy()
                            pt_view["State"] = pt_view["State"].map(status_map).fillna(pt_view["State"])
                            # 定投表达式翻译为中文展示
                            pt_view["Product Targeting Expression"] = pt_view["Product Targeting Expression"].map(target_expr_map).fillna(pt_view["Product Targeting Expression"])
                            pt_view["Click-through Rate"] = pt_view["Click-through Rate"].apply(format_percent)
                            pt_view["Conversion Rate"] = pt_view["Conversion Rate"].apply(format_percent)
                            pt_view["ACOS"] = pt_view["ACOS"].apply(format_percent)
                            pt_view.columns = [
                                "投放状态","定位ASIN","出价","曝光","点击","点击率","花费($)","销售额($)","订单量","销量","转化率","ACOS","CPC","ROAS"
                            ]
                            cfg_pt = {c: st.column_config.Column(width="stretch") for c in pt_view.columns}
                            pt_click = st.dataframe(
                                pt_view,
                                use_container_width=True,
                                height=180,
                                on_select="rerun",
                                selection_mode="single-row",
                                key=f"pt_{ag_name}",
                                column_config=cfg_pt
                            )
                            sel_tgt_rows = pt_click.selection.rows
                            if len(sel_tgt_rows) > 0:
                                # 清空关键词缓存
                                st.session_state.sel_keyword_text = ""
                                st.session_state.sel_match_type_raw = ""
                                # 从原始pt表取未翻译英文原值，用于搜索词匹配
                                st.session_state.sel_target_expr = pt.iloc[sel_tgt_rows[0]]["Product Targeting Expression"]
                    # 投放关键词（核心修改：选中时存储三重匹配字段）
                    with st.container():
                        kw = ag_data[ag_data["Entity"] == "Keyword"]
                        if len(kw) > 0:
                            st.markdown("<div class='block-title'>▸ 投放关键词（单击行匹配下方客户搜索词）</div>", unsafe_allow_html=True)
                            kw_cols = ["State","Keyword Text","Match Type","Bid","Impressions","Clicks","Click-through Rate","Spend","Sales","Orders","Units","Conversion Rate","ACOS","CPC","ROAS"]
                            kw_view = kw[kw_cols].copy()
                            kw_view["State"] = kw_view["State"].map(status_map).fillna(kw_view["State"])
                            kw_view["Match Type CN"] = kw_view["Match Type"].map(match_type_map)
                            kw_view["Click-through Rate"] = kw_view["Click-through Rate"].apply(format_percent)
                            kw_view["Conversion Rate"] = kw_view["Conversion Rate"].apply(format_percent)
                            kw_view["ACOS"] = kw_view["ACOS"].apply(format_percent)
                            kw_display_cols = [
                                "State","Keyword Text","Match Type CN","Bid","Impressions","Clicks","Click-through Rate","Spend","Sales","Orders","Units","Conversion Rate","ACOS","CPC","ROAS"
                            ]
                            kw_display = kw_view[kw_display_cols].copy()
                            kw_display.columns = [
                                "投放状态","关键词文本","匹配方式","出价","曝光","点击","点击率","花费($)","销售额($)","订单量","销量","转化率","ACOS","CPC","ROAS"
                            ]
                            cfg_kw = {c: st.column_config.Column(width="stretch") for c in kw_display.columns}
                            kw_select = st.dataframe(
                                kw_display,
                                use_container_width=True,
                                height=390,
                                on_select="rerun",
                                selection_mode="single-row",
                                key=f"kw_{ag_name}",
                                column_config=cfg_kw
                            )
                            sel_kw_rows = kw_select.selection.rows
                            if len(sel_kw_rows) > 0:
                                sel_kw_idx = sel_kw_rows[0]
                                # 存储三重匹配条件：广告活动、关键词原文、原始匹配类型英文
                                st.session_state.sel_target_expr = ""
                                st.session_state.sel_keyword_text = kw_view.iloc[sel_kw_idx]["Keyword Text"]
                                st.session_state.sel_match_type_raw = kw_view.iloc[sel_kw_idx]["Match Type"]
                    # 否定ASIN（默认折叠，无数据自动隐藏）
                    neg_pt = ag_data[ag_data["Entity"] == "Negative Product Targeting"]
                    if len(neg_pt) > 0:
                        with st.expander("▸ 否定ASIN", expanded=False):
                            neg_pt_cols = ["State","Product Targeting Expression","Impressions","Clicks","Click-through Rate","Spend","Sales","Orders","Units","Conversion Rate","ACOS","CPC","ROAS"]
                            neg_pt_view = neg_pt[neg_pt_cols].copy()
                            neg_pt_view["State"] = neg_pt_view["State"].map(status_map).fillna(neg_pt_view["State"])
                            neg_pt_view["Click-through Rate"] = neg_pt_view["Click-through Rate"].apply(format_percent)
                            neg_pt_view["Conversion Rate"] = neg_pt_view["Conversion Rate"].apply(format_percent)
                            neg_pt_view["ACOS"] = neg_pt_view["ACOS"].apply(format_percent)
                            neg_pt_view.columns = [
                                "投放状态","否定ASIN","曝光","点击","点击率","花费($)","销售额($)","订单量","销量","转化率","ACOS","CPC","ROAS"
                            ]
                            cfg_neg_tgt = {c: st.column_config.Column(width="stretch") for c in neg_pt_view.columns}
                            st.dataframe(neg_pt_view, use_container_width=True, height=180, column_config=cfg_neg_tgt)
                    # 否定关键词（默认折叠，无数据自动隐藏）
                    neg_kw = ag_data[ag_data["Entity"] == "Negative Keyword"]
                    if len(neg_kw) > 0:
                        with st.expander("▸ 否定关键词", expanded=False):
                            neg_kw_cols = ["State","Keyword Text","Match Type","Impressions","Clicks","Click-through Rate","Spend","Sales","Orders","Units","Conversion Rate","ACOS","CPC","ROAS"]
                            neg_kw_view = neg_kw[neg_kw_cols].copy()
                            neg_kw_view["State"] = neg_kw_view["State"].map(status_map).fillna(neg_kw_view["State"])
                            neg_kw_view["Match Type"] = neg_kw_view["Match Type"].map(match_type_map).fillna(neg_kw_view["Match Type"])
                            neg_kw_view["Click-through Rate"] = neg_kw_view["Click-through Rate"].apply(format_percent)
                            neg_kw_view["Conversion Rate"] = neg_kw_view["Conversion Rate"].apply(format_percent)
                            neg_kw_view["ACOS"] = neg_kw_view["ACOS"].apply(format_percent)
                            neg_kw_view.columns = [
                                "投放状态","否定关键词","匹配方式","曝光","点击","点击率","花费($)","销售额($)","订单量","销量","转化率","ACOS","CPC","ROAS"
                            ]
                            cfg_neg_kw = {c: st.column_config.Column(width="stretch") for c in neg_kw_view.columns}
                            st.dataframe(neg_kw_view, use_container_width=True, height=180, column_config=cfg_neg_kw)
                st.divider()
        else:
            st.info("该广告无细分广告组数据")
    st.divider()

    # ===================== 五、客户搜索词明细【修正：定投双条件同时过滤 + 定投表达式中文翻译】 =====================
    with st.container():
        st.subheader("五、客户搜索词明细（匹配上方关键词/ASIN自动过滤）")
        user_search_input = st.text_input("手动搜索客户词")
        if df_search is not None:
            st_raw = df_search.copy()
            kw_text = st.session_state.sel_keyword_text
            match_raw = st.session_state.sel_match_type_raw
            target_expr = st.session_state.sel_target_expr
            camp_name = sel_camp_name

            # 分支匹配逻辑
            if kw_text != "" and match_raw != "":
                # 关键词三重严格匹配
                if "Campaign Name (Informational only)" in st_raw.columns:
                    st_raw = st_raw[st_raw["Campaign Name (Informational only)"] == camp_name]
                if "Keyword Text" in st_raw.columns and "Match Type" in st_raw.columns:
                    st_raw = st_raw[(st_raw["Keyword Text"] == kw_text) & (st_raw["Match Type"] == match_raw)]
            elif target_expr != "":
                # ASIN定投双重精确匹配（完全按你的规则）
                # 第一层：广告活动完全相等
                if "Campaign Name (Informational only)" in st_raw.columns:
                    st_raw = st_raw[st_raw["Campaign Name (Informational only)"] == camp_name]
                # 第二层：定投表达式完全相等
                if "Product Targeting Expression" in st_raw.columns:
                    st_raw = st_raw[st_raw["Product Targeting Expression"] == target_expr]

            # 手动文本检索（模糊包含）
            if user_search_input.strip() != "":
                search_col = "Customer Search Term" if "Customer Search Term" in st_raw.columns else "Search Term"
                st_raw = st_raw[st_raw[search_col].str.contains(user_search_input, na=False, case=False)]
            # 花费降序
            if len(st_raw) > 0 and "Spend" in st_raw.columns:
                st_raw = st_raw.sort_values("Spend", ascending=False, ignore_index=True)
            # 展示字段
            origin_search_col = "Customer Search Term" if "Customer Search Term" in st_raw.columns else "Search Term"
            display_cols = [
                origin_search_col,
                "Match Type","Keyword Text","Product Targeting Expression","Impressions","Clicks",
                "Click-through Rate","Spend","Sales","Orders","Units",
                "Conversion Rate","ACOS","CPC","ROAS"
            ]
            # 只保留报表存在的列，防止报错
            display_cols = [c for c in display_cols if c in st_raw.columns]
            show_df = st_raw[display_cols].copy()
            # 格式化
            if "Match Type" in show_df.columns:
                show_df["Match Type"] = show_df["Match Type"].map(match_type_map).fillna("无")
            # 新增：定投表达式中英转换
            if "Product Targeting Expression" in show_df.columns:
                show_df["Product Targeting Expression"] = show_df["Product Targeting Expression"].map(target_expr_map).fillna(show_df["Product Targeting Expression"])
            pct_fields = ["Click-through Rate","Conversion Rate","ACOS"]
            for f in pct_fields:
                if f in show_df.columns:
                    show_df[f] = show_df[f].apply(format_percent)
            # 中文表头映射
            name_map = {
                origin_search_col:"客户搜索词",
                "Match Type":"匹配方式",
                "Keyword Text":"投放关键词",
                "Product Targeting Expression":"定投表达式",
                "Impressions":"曝光",
                "Clicks":"点击",
                "Click-through Rate":"点击率",
                "Spend":"花费($)",
                "Sales":"销售额($)",
                "Orders":"订单量",
                "Units":"Units",
                "Conversion Rate":"转化率",
                "ACOS":"ACOS",
                "CPC":"CPC",
                "ROAS":"ROAS"
            }
            new_header = [name_map[c] for c in display_cols]
            show_df.columns = new_header
            cfg_search = {c: st.column_config.Column(width="stretch") for c in show_df.columns}
            if len(show_df) > 0:
                st.dataframe(show_df, use_container_width=True, height=390, column_config=cfg_search)
            else:
                st.warning("无匹配的客户搜索词数据")
        else:
            st.info("未上传SP Search Term报表文件")
