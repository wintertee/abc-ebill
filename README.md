# abc-ebill

农行（ABC）金穗信用卡**电子对账单解析器**：把银行发的 `.eml` 对账单解析成多份 xlsx（账务说明 / 交易明细 / 按卡汇总 / 每卡明细），并自动**核对**交易明细与账务说明是否一致，不一致直接报错。

> 支持两种用法：命令行（CLI）和浏览器网页版（Pyodide，数据不出本机）。

## 目录结构

```
.
├── main.py            # CLI 入口
├── web_adapter.py     # 网页版入口（Pyodide 用）
├── index.html         # 网页版页面
├── pyproject.toml     # uv 项目（依赖 openpyxl / wcwidth）
├── .gitignore         # 忽略 bills/（含隐私数据）与 Python 缓存
└── bills/             # 原始 .eml + 每月输出（已 gitignore）
    ├── 2608.eml
    └── 2608/
        ├── 2608_summary.xlsx
        ├── 2608_transactions.xlsx
        ├── 2608_cards.xlsx
        └── 2608_card_0921.xlsx …   # 每张卡一份明细
```

## 安装 / 环境

依赖 [uv](https://docs.astral.sh/uv/)（或 `pip install uv`）。首次运行 `uv run` 会自动创建 `.venv` 并安装依赖。

```bash
uv sync
```

## 用法一：命令行

按"期号"处理（账单放 `bills/<期号>.eml`，输出自动进 `bills/<期号>/`）：

```bash
# 放账单
cp xxx.eml bills/2608.eml

# 处理（读取 bills/2608.eml，输出到 bills/2608/）
uv run main.py 2608
```

也可以直接给 `.eml` 路径（输出到它旁边）：

```bash
uv run main.py bills/2608/2608.eml
```

`--outdir` 指定输出目录，`--prefix` 覆盖文件名前缀。

### 输出文件

| 文件 | 内容 |
|---|---|
| `2608_summary.xlsx` | 账务说明 + 邮件头信息 |
| `2608_transactions.xlsx` | 全部交易明细（金额按币种拆列） |
| `2608_cards.xlsx` | 按卡汇总（含合计行、还款日） |
| `2608_card_XXXX.xlsx` | 每张卡的明细，顶部为该卡汇总 |

核对不通过时脚本退出码非 0，并在 stderr 说明哪一项对不上。

## 用法二：网页版（Pyodide）

处理完全在浏览器本地完成，账单数据**不会上传**到任何服务器。

```bash
uv run python -m http.server 8765
```

浏览器打开 `http://localhost:8765/`，选择 `.eml` 文件 → 点"转换"：

- 页面直接显示核对报告和按卡汇总
- 一个"下载全部"按钮，打包下载 `2608_对账单.zip`（内含全部 xlsx）

> 首次加载需联网拉取 Pyodide 运行时（约几 MB，浏览器会缓存）；依赖 openpyxl/wcwidth 通过 micropip 从 PyPI 加载。

本地快速验证网页模块（不启动浏览器）：

```bash
uv run python web_adapter.py bills/2608.eml
```

## 核对逻辑

对账单的账务说明与交易明细必须闭合，否则 `ValueError`：

```
本期账单金额        == 消费 + 费用
本期还款、退货金额  == 还款 + 退货
本期应还 − 本期溢缴 == 上期应还 − 上期溢缴 + 本期账单 − 本期还款退货 − 本期调整
```

另外按卡分组的消费/退货/还款合计必须等于全账户合计。

## 按卡"待还款总额"口径

```
待还款总额(本期) = 消费(含费用) − 退货
```

**还款（本期入账的还款）属于对上一账单周期的还款**，不计入本期待还款，单独列在"还款(上期)"。

## 隐私

`bills/` 目录里的 `.eml` 和 xlsx 包含卡号、姓名、完整交易记录等隐私数据，已加入 `.gitignore`。提交仓库时只提交代码文件即可。
