---
name: fill-mabang
description: 读取桌面"马帮信息.xlsx"，使用Playwright自动填写马帮ERP订单表单。当用户说"填马帮"时自动调用。
allowed-tools: Bash(python *) Bash(taskkill*) Bash(rm *)
---

# 填写马帮ERP订单表单

读取桌面Excel数据，自动填写马帮ERP的新建订单表单。

## 执行步骤

运行脚本完成填写：

```
python <skill-directory>/scripts/mabang_fill.py
```

脚本会：
1. 读取 `~/Desktop/马帮信息.xlsx` 中的订单数据
2. 启动Playwright浏览器（使用持久化session保持登录）
3. 导航到马帮ERP新建订单页面
4. 在iframe中自动填写所有字段，包括：
   - 文本字段：收件人、电话、邮箱、地址、城市、省/州、邮编、交易号、内部单号、客户ID
   - 国家下拉（xm-select组件）：自动搜索并选择对应国家
   - 店铺下拉：自动选择"售后(Global区域)"
   - 业务员：选完店铺后自动填入"赖丹"
   - 自定义分类：自动勾选"售后(补发/换货...)"
   - 付款时间：自动填入当天日期时间（UTC+8）
5. 输出填写结果

## 注意事项

- 如果浏览器提示未登录，告诉用户在浏览器中手动登录，脚本会自动等待并继续
- 订单号(platformOrderId)保持系统默认值，不要填写
- 填完后直接告诉用户"填好了"，不需要截图确认
- 浏览器保持打开，剩下的交给用户操作，无需监控
