# 经营者集中申报审查 Skill

面向律师、法务和交易团队的中国经营者集中初筛。Skill 根据用户提供的交易结构、经营者控制权、集团关系、全球及中国境内营业额、交易阶段和附件材料，生成可追溯的门槛、程序、材料及抢跑风险审查包。

## 使用方式

先准备交易结构。推荐 JSON：

```json
{
  "title": "甲公司收购乙公司股权",
  "transaction_type": "取得控制权",
  "closing_status": "已签约未交割",
  "operators": [
    {"operator_id": "A", "name": "甲公司", "role": "收购方"},
    {"operator_id": "B", "name": "乙公司", "role": "目标经营者"}
  ]
}
```

准备阶段会记录输入 SHA-256，并把附件逐份转换为临时 Markdown。云端运行必须明确确认隐私：

```bash
python3 scripts/run_pipeline.py prepare \
  --transaction transaction.json \
  --attachment 交易协议.docx \
  --workspace TEMP_DIR \
  --processing-environment cloud \
  --privacy-confirmed
```

Agent 读取 `TEMP_DIR/merger-bundle.json` 和 Markdown 索引后，提交 `analysis.json`。随后生成报告：

```bash
python3 scripts/run_pipeline.py report \
  --bundle TEMP_DIR/merger-bundle.json \
  --analysis-json analysis.json \
  --output-dir OUTPUT_DIR \
  --cleanup
```

## 输出

输出目录包含：

- `<主题>_经营者集中申报审查.xlsx`：审查摘要、经营者与交易、控制权分析、营业额记录、程序分流、申报材料清单、风险与建议、法律依据、待核验和处理记录。
- `<主题>_经营者集中申报审查.html`：可离线打开的风险与建议概览。

金额默认按亿元记录，但每条营业额必须同时提供年度、币种、范围和来源定位。缺少营业额、控制权或交易阶段资料时，只能输出“待核验”，不能推定未达到申报门槛。

## 边界

本 Skill 只基于用户主动提供的材料，不联网补充企业关联关系，不代表用户向市场监管总局提交申报，不判断最终申报义务或审查结果，也不替代反垄断律师和律所内部审批。输出可能包含商业秘密，应限制访问并妥善保管。

## 验证

```bash
python3 tests/test_pipeline.py
python3 scripts/check_pipeline.py .
```
