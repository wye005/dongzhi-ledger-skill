# -*- coding: utf-8 -*-
"""OA 订单评审明细 CSV → 合约订单评审表YYYY-MM-DD.xlsx（沿用历史 66 列格式）"""
import csv
import sys

import openpyxl

SRC = r"E:\workbuddy-temp\tmp\oa_review_detail.csv"
DST = sys.argv[1]

# 历史 66 列顺序 ← 接口字段
MAP = [
    ("OA流水号", "SN"), ("评审状态", "StatusStr"), ("制单人", "Applicant"), ("制单日期", "App_Date"),
    ("销售性质", "DDXZ"), ("BOSA方案", "BOSA"), ("订单分类", "oa_vtrantype"), ("订单类型", "OrderType"),
    ("公司订单号", "erp_vbillcode"), ("客户编码", "erp_ccustomerid"), ("客户名称", "oa_customer_name"),
    ("客户分类", "oa_customer_type"), ("执行框架", "erp_vdef20"), ("报价编号", "oa_erp_bjprice"),
    ("客户订单号", "erp_vdef2"), ("客户下单日期", "CustomerOrderDate"), ("客户要求交期", "oa_require_delivery"),
    ("评审日期", "oa_comfirm_delivery"), ("评审跟进人", "erp_creator"), ("销售人员", "oa_cbiz_name"),
    ("销售部门", "oa_cbiz_dept"), ("销售区域", "SalesArea"), ("交货方式", "DeliveryType"),
    ("其他交货方式", "OtherDeliveryType"), ("运输方式", "erp_ctransporttype"), ("订单文本形式", "OrderTextForm"),
    ("供料方式", "FeedingWay"), ("客户验货要求", "InspectionRequire"), ("客供物料交货方式", "MaterialDeliveryType"),
    ("付款条件", "Payment"), ("收款协议", "oa_erp_cpaytermid"), ("币种", "erp_corigcurrency"),
    ("汇率", "ExchangeRate"), ("税率", "TaxRate"), ("订单总数量", "erp_ntotalnum"), ("单位", "UnitCode"),
    ("物料编码", "erp_cmaterialvid"), ("物料名称", "oa_cmaterial"), ("是否赠品", "erp_flargess"),
    ("产品类别", "ProductCategory"), ("客户机型", "CustomerModel"), ("公司机型", "CompanyModel"),
    ("客户物料代码", "CustomerMaterialID"), ("客户指令PKG", "khzlhPKG"), ("订单行数量", "OrderQuantity"),
    ("不含税单价", "NoInTaxPrice"), ("含税单价", "InTaxPrice"), ("不含税金额", "NoInTaxTotal"),
    ("含税金额", "InTaxTotal"), ("项目名称局点", "ProjectName"),
]
NUMS = {"ExchangeRate", "TaxRate", "erp_ntotalnum", "OrderQuantity", "NoInTaxPrice",
        "InTaxPrice", "NoInTaxTotal", "InTaxTotal", "erp_flargess"}

rows = list(csv.DictReader(open(SRC, encoding="utf-8-sig")))
wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Sheet1"
ws.cell(1, 1, "合约、订单审批流程")
for j, (cn, _) in enumerate(MAP, start=1):
    ws.cell(2, j, cn)
# 保持历史文件的 66 列宽度（末尾 16 列空）
for j in range(len(MAP) + 1, 67):
    ws.cell(2, j, None)

for i, r in enumerate(rows, start=3):
    for j, (cn, key) in enumerate(MAP, start=1):
        v = r.get(key, "")
        if key in NUMS:
            try:
                v = float(v) if v not in ("", None) else None
            except (TypeError, ValueError):
                v = None
        ws.cell(i, j, v)

wb.save(DST)
print("写出", DST, "数据行", len(rows))
