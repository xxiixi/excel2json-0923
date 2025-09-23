# Excel2JSON for Mac

这是 Excel2JSON 工具的 Mac 版本，使用 Node.js 实现，无需安装 Microsoft Excel。

## 安装和运行

### 1. 确保已安装 Node.js
```bash
node --version
```
如果没有安装，请访问 [Node.js 官网](https://nodejs.org/) 下载安装。

### 2. 安装依赖
```bash
npm install
```

### 3. 运行转换
```bash
# 转换当前目录下所有 Excel 文件（简单模式）
./run.sh

# 转换指定文件（简单模式）
./run.sh 文件1.xlsx 文件2.xlsx

# 使用高级模式（需要特殊标记）
./run.sh advanced

# 转换指定文件到指定输出目录
./run.sh 文件1.xlsx 文件2.xlsx 输出目录
```

## 转换模式

### 简单模式（默认）
- 自动识别 Excel 表格结构
- 将 ID 列作为对象的键
- 适合大多数普通 Excel 文件
- 无需特殊标记

## 输出格式

转换后的 JSON 文件会保存在 `output` 目录中，格式如下：

```json
{
  "工作表名": {
    "ID1": {
      "列名1": "值1",
      "列名2": "值2"
    },
    "ID2": {
      "列名1": "值1",
      "列名2": "值2"
    }
  }
}
```

## 配置文件

可以创建 `Excel2Json.config.js` 文件来自定义配置：

```javascript
var g_sourceFolder = "/path/to/source";
var g_targetFolder = "output";
var g_prettyOutput = true;
```

## 故障排除

1. **如果转换结果为空**：检查 Excel 文件是否有数据，或尝试使用简单模式
2. **如果出现编码问题**：确保 Excel 文件保存为 UTF-8 编码
3. **如果依赖安装失败**：尝试使用 `npm install --force`

## 示例

转换你的 Excel 文件：
```bash
./run.sh 深夜模式局内词条.xlsx 深夜模式局外词条.xlsx
```

输出文件将保存在 `output/` 目录中。
