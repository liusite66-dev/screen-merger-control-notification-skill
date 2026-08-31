---
name: screen-merger-control-notification-skill
activation: /screen-merger-control-notification-skill
license: MIT
metadata:
  author: liusite66-dev
  version: 1.0.0
  created: 2026-08-31
  last_reviewed: 2026-08-31
  review_interval_days: 180
provenance:
  maintainer: liusite66-dev
description: >-
  Screen a user-provided China merger-control transaction for control, filing
  thresholds, procedure routing, materials and gun-jumping risks. Use transaction
  JSON/XLSX and optional attachments; produce traceable Chinese XLSX and offline
  HTML. Do not make a final filing opinion, invent turnover or control, search
  company data, submit to SAMR, or predict review outcome.
---
# 经营者集中申报审查

先取得云端隐私确认并建议脱敏。材料只读，记录 SHA-256、来源和缺口。先运行
`prepare`，由 Agent 按索引/材料定位补充语义判断，再将结构化 analysis JSON 交给
`report`。确定性脚本计算营业额门槛、集团内重组、程序候选和抢跑风险；缺资料只能是
“待核验”，不得把未提供当作未达到门槛。

输出为 `<主题>_经营者集中申报审查.xlsx` 与离线 HTML；不代表市场监管总局决定或正式法律意见。

## Gotchas

- 法定门槛会随法律和政策变化；脚本中的阈值只是当前配置快照，报告必须记录版本并由人工核验。
- 缺失营业额、控制权或交易阶段资料只能输出“待核验”。
- 交易文件可能含商业秘密和个人信息，限制访问并妥善保管输出。
