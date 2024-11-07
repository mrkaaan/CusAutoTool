import pandas as pd

# 示例数据框
data = {
    "原始单号": [
        "123456",
        "'123456",
        "789012-34",
        "\"345678",
        "=901234",
        "“567890”",
        None
    ]
}

df = pd.DataFrame(data)

# 清理“原始单号”列
df["原始单号"] = (
    df["原始单号"]
    .astype(str)  # 转为字符串类型
    .str.replace(r"-\d+$", "", regex=True)  # 去除斜杠及其后面的数字
    .str.replace(r"^[\'\"]", "", regex=True)  # 去掉以单引号或双引号开头的字符
    .str.replace(r"[=“”\"\'']", "", regex=True)  # 去除指定符号
    .apply(lambda x: str(int(x)) if x.isdigit() else x)  # 确保只包含数字，去掉无效字符
)

# 打印结果
print(df)
