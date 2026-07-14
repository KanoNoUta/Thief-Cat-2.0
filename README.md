# Thief Cat 2.0

Thief Cat 2.0 是一个 Databricks GLM 5.2 Access Token 自动化工具，负责邮箱验证、注册或登录、Workspace 进入，以及 AI Gateway Token 生成和账号分类。

## 特性

- 仅支持 Outlook Graph OAuth 邮箱，流程更稳定、可重复登录。
- Outlook pending/success 队列自动迁移。
- 动态等待 AI Gateway 和 GLM 5.2，不依赖固定模型 URL。
- 一键点击 `Generate Access Token`，无需填写 PAT 表单。
- 只有检测到真实二维码时才等待人工操作。
- 成功输出固定为 `Domain,Token`，同一 Domain 自动更新。

## 工作流

1. 通过 Outlook Graph 获取验证码并登录 Databricks。
2. 处理 account setup，进入 Workspace。
3. 打开 AI Gateway，等待模型列表加载。
4. 点击 `GLM 5.2` 和 `Generate Access Token`。
5. 将结果写入 `glm_keys.csv`。
6. Outlook 账号成功后，从 pending 文件移动到 success 文件。

只有检测到真实二维码时才等待人工扫码。其他未知页面会截图并明确失败，不会要求按 Enter 或手动粘贴 Token。

## 文件

- `get_glm_key.py`：发布版主入口。
- `auto_register.py`：完整注册、登录、Workspace 和 Token 流程。
- `outlook_graph.py`：Microsoft OAuth 与 Graph 邮件读取。
- `outlook_pending.txt`：尚未成功的 Outlook 凭据队列。
- `outlook_success.txt`：已经成功的 Outlook 凭据。
- `glm_keys.csv`：成功结果，固定为 `Domain,Token` 两列。
- `registered_accounts.csv`：运行审计记录，包含账号、状态和错误。

以上 TXT/CSV 都包含敏感数据，已经由 `.gitignore` 排除。

## 安装

```powershell
python -m pip install -r .\requirements.txt
```

默认使用本机 Chrome，也可以指定 Edge：

```powershell
python .\get_glm_key.py --browser-channel msedge
```

## Outlook 队列

在 `outlook_pending.txt` 中每行放一个账号：

```text
email----password----refresh_token----client_id
```

可以从不含真实凭据的 `outlook_pending.example` 开始配置。

也兼容：

```text
email----password----client_id----refresh_token
```

未指定 `--email` 时自动取 pending 文件第一条。只有完整生成 Token 后才会将原始凭据行移动到 `outlook_success.txt`；失败账号会留在 pending 文件中。

测试邮箱读取：

```powershell
python .\get_glm_key.py --mail-test
```

新账号注册：

```powershell
python .\get_glm_key.py `
  --mail-timeout 300
```

已注册账号续跑：

```powershell
python .\get_glm_key.py `
  --resume `
  --workspace "https://your-workspace.cloud.databricks.com" `
  --mail-timeout 300
```

覆盖默认分类文件：

```powershell
python .\get_glm_key.py `
  --outlook-credential-file .\custom_pending.txt `
  --outlook-success-file .\custom_success.txt
```

## 输出格式

`glm_keys.csv`：

```csv
Domain,Token
https://dbc-example.cloud.databricks.com,dapi...
```

同一 Domain 再次成功时会更新 Token，不会追加重复 Domain。

## 验证

```powershell
python -m py_compile .\get_glm_key.py .\auto_register.py .\outlook_graph.py .\mail_common.py
python -m unittest discover -s .\tests -v
python .\get_glm_key.py --help
```
