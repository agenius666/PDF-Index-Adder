# PDF Index Adder | PDF索引号添加工具

## 🇨🇳 中文使用说明
### 📥 安装步骤
```bash
pip install -r requirements.txt
```
字体配置：将SimSun.ttc放在：
程序根目录
或系统字体目录（C:\Windows\Fonts 或 /Library/Fonts）  

📝 Excel模板示例
| PDF文件路径          | 索引号         |
|----------------------|---------------|
| files/报告.pdf       | FIN-2023-001  |
| ../财务/年报.pdf     | ACCT-2023-Q4  |  

🖱️ 界面操作  
1，点击"浏览" → 选择Excel文件  
2，指定输出目录 → 自动创建文件夹  
3，点击"开始处理" → 查看实时进度条  
4，错误日志 → 输出目录下的error.log  

⚙️ 技术参数  
1，配置项	默认值  
2，字体大小	页面高度的3%  
3，右侧边距	页面宽度的3%  
4，输出命名规则	{索引号}_{原文件名}  

## 🇺🇸 English Documentation

📥 Installation
```bash
pip install -r requirements.txt
```

Font setup: Place SimSun.ttc in:
Program root directory
OR system fonts folder (C:\Windows\Fonts or /Library/Fonts)

📝 Excel Template
| PDF Path             | Index Number   |
|----------------------|---------------|
| files/report.pdf    | FIN-2023-001  |
| ../finance/annual.pdf | ACCT-2023-Q4 |

🖱️ GUI Operations  
1，Click "Browse" → Select Excel file  
2，Choose output dir → Auto-create folders  
3，Click "Start" → Monitor progress bar  
4，Error logs → error.log in output dir  

⚙️ Technical Specs  
1，Configuration	Default Value  
2，Font Size	3% of page height  
3，Right Margin	3% of page width  
4，Output Naming	{Index}_{OriginalName}  
