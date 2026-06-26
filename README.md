# PRD 测试用例生成工具

从 PRD、会议纪要、表结构、分区信息和开发代码中提炼数据测试需求，
经过待确认点收敛后，自动生成最终版需求提炼、测试用例和 SQL 校验脚本。

## 在线使用

部署后访问链接，在左侧侧边栏填写你自己的 API Key、Base URL、Model 即可使用。

## 本地运行

```bash
pip install -r requirements.txt
streamlit run app.py
```

## 功能流程

1. 上传或粘贴 PRD
2. AI 生成初版需求提炼和待确认点
3. 用户补充待确认点说明（支持多轮收敛）
4. AI 生成最终版需求提炼表
5. AI 基于最终版需求提炼表生成测试用例和 SQL
