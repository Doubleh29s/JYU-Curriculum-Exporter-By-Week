# 📘 JYU Curriculum Exporter By Week

一个使用学校VPN登录 **嘉应学院教务系统** 并导出个人课表到 **Excel 文件** 的脚本工具。 
通过 VPN 登录校园系统，自动获取课表数据并生成格式化的 `课程表.xlsx` 文件。
使用的入口是教务系统里面，按周次的学生课表查询（课表内容每周更新，若有老师换课，课表内容也可以同步）

## ✨ 功能特点

- 🔑 自动模拟登录 VPN 和教务系统
- 📥 获取课表 JSON 数据并提取核心字段
- 📊 自动生成 **7 天 × 12 节课** 的课程表
- 📂 导出到 Excel 文件（支持换行显示、边框分隔）

## 📂 项目结构

```
.
├── get_curriculum.py      # 主程序，负责登录、获取课表并调用导出模块
├── export_curriculum.py   # 导出模块，将课表写入 Excel 文件
├── requirements.txt       # 依赖文件
├── user_config.json       # 用户配置文件 （需填写账号密码）
└── README.md              # 项目说明
```

## 📦 安装依赖方式

### 方式一：使用虚拟环境（推荐）

```bash
# 创建虚拟环境
python -m venv venv

# 激活虚拟环境
# Windows
venv\Scripts\activate
# Linux / macOS
source venv/bin/activate

# 安装依赖
pip install -r requirements.txt
```

### 方式二：直接安装

```bash
pip install -r requirements.txt
```


## ⚙️ 配置方法

在运行前，请先修改 `user_config.json`，填入你的账号和密码，按照标注修改，其他的不需要修改

```json
{
  "vpn_credentials": {
    "username": "引号里填写你的VPN学号",
    "password": "引号里填写你的VPN密码"
  },
  "system_credentials": {
    "csrftoken": "",
    "language": "zh_CN",
    "yhm": "引号里填写你的教务系统学号",
    "mm": "引号里填写你的教务系统密码"
  },
  "curriculum_param": {
    "xnm": "",
    "xqm": "",
    "zs": "",
    "kblx": 1,
    "doType": "app"
  },
  "vpn_login_url": "https://ids-jyu-edu-cn.webvpn.jyu.cn/authserver/login",
  "public_key_url": "https://jwcjwxt-jyu-edu-cn-443.webvpn.jyu.cn/xtgl/login_getPublicKey.html",
  "system_login_url": "https://jwcjwxt-jyu-edu-cn-443.webvpn.jyu.cn/xtgl/login_slogin.html",
  "curriculum_url": "https://jwcjwxt-jyu-edu-cn-443.webvpn.jyu.cn/kbcx/xskbcxMobile_cxXsKb.html?gnmkdm=N2154",
  "curriculum_html_url": "https://jwcjwxt-jyu-edu-cn-443.webvpn.jyu.cn/kbcx/xskbcxZccx_cxXskbcxIndex.html?gnmkdm=N2154&layout=default"
}
```


## ▶️ 使用方法

在项目根目录下运行：

```bash
python get_curriculum.py
```

运行成功后，将会在目录下生成一个 `课程表.xlsx` 文件。

## ⚠️ 注意事项

- 依赖vpn和教务系统，如果学校修改，该脚本就无法使用。
- 脚本是依靠教务系统的周次课表，有的课程在特定周，尽量在新的一周运行一次脚本，防止忘记上课😅
- 导出的 `课程表.xlsx` 默认保存在项目根目录，可在代码中修改保存路径。

## 之后改善
- 后面可能会构造一个本地的app来运行脚本或者微信小程序