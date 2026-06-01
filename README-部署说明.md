# JTT电商AI助手 — Railway 部署说明

## 步骤 1：上传代码到 GitHub

```bash
# 1.1 在项目目录初始化 git
cd "/Users/admin/bb plan1"
git init

# 1.2 添加所有文件
git add .

# 1.3 提交
git commit -m "first commit"

# 1.4 在 GitHub 上创建仓库
#     浏览器打开 https://github.com/new
#     仓库名：jtt-assistant（或任意名字）
#     创建后复制仓库地址: https://github.com/你的用户名/jtt-assistant.git

# 1.5 推送代码到 GitHub
git remote add origin https://github.com/你的用户名/jtt-assistant.git
git branch -M main
git push -u origin main
```

## 步骤 2：在 Railway 部署

### 2.1 注册 Railway
1. 打开 https://railway.app
2. 点击 **Login with GitHub** 用 GitHub 账号登录

### 2.2 创建项目
1. 点击 **New Project** → **Deploy from GitHub repo**
2. 选择你刚创建的 `jtt-assistant` 仓库
3. Railway 会自动检测 `requirements.txt` 和 `Procfile`

### 2.3 部署
1. 点击 **Deploy**
2. 等待几分钟，Railway 会自动安装依赖并启动
3. 部署完成后，点击 **Generate Domain** 生成公网域名
4. 你会得到一个 `https://jtt-assistant.up.railway.app` 格式的地址

### 2.4 访问
```
https://你的域名.up.railway.app       → 账单生成
https://你的域名.up.railway.app/invoice → 发票转换
```

## 文件结构说明

```
bb plan1/                   ← git 仓库根目录
├── Procfile                ← Railway 启动命令
├── requirements.txt        ← Python 依赖
├── .gitignore
├── README-部署说明.md
├── TR账单自动生成/
│   ├── app.py              ← Flask 主程序
│   ├── gen_bill.py         ← 账单生成引擎
│   ├── 账单模板.xlsx        ← 账单模板文件
│   └── templates/          ← HTML 页面
└── 发票转换/
    ├── convert_invoice.py  ← 发票转换引擎
    ├── 天图下单发票.xlsx     ← 天图模板
    └── 航乐*.xls           ← 航乐模板
```

## Railway 免费额度

| 项目 | 额度 |
|------|------|
| 运行时间 | 每月 500 小时 |
| 带宽 | 每月 100 GB |
| 项目数 | 3 个 |
| 域名 | 自动生成 `.up.railway.app` |

> 500 小时 ≈ 每天 16 小时，工作日够用。如果 24 小时跑，需要付费 ¥5-10/月。
