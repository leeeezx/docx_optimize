# READMe

**本skill为对`anthropics/skills.git/skills/docx`的优化版本**

**主要优化点：**
1. 通过添加word宏命令与py脚本的结合，解决了原版本中`mathtype`公式无法转换的问题。


## 环境设置

1. 确保已下载`mathtype`，并且word中存在相关宏命令`MTCommand_ConvertEqns`。
    - `mathtype`设置如图所示，[fig1](fig\mathtype_option01.png)、[fig2](fig\mathtype_option02.png)
2. Windows 系统。
3. 本机安装 Microsoft Word。
4. Python 3
5. 添加`scripts/math_ops/公式转换MathML至OMML.bas`至本机word的Normal模板中.



### 自查方法

- MTCommand_ConvertEqns
  - 打开word - 调出顶部任务栏的`开发工具` - 找到`开发工具`中的`宏`- 点开即可查看当前存在的宏命令 
- `公式转换MathML至OMML.bas`添加
  - 随便打开一个word文档 - 按下Alt + F11 -  点击左侧的`Normal` - 点击上方的`文件` - 点击`导入文件` - 选择`scripts/math_ops/公式转换MathML至OMML.bas` - 导入成功后即可在左侧看到`Module1`，点击即可查看代码。
  [如图所示](fig\PixPin_2026-03-01_15-48-13.png)