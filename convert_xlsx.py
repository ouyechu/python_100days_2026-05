import os
import unicodedata

import pandas as pd

# ===================== 【只需确认这一行】 =====================
# 你的文件夹（所有 xlsx 都在这里）
FOLDER = r"E:\DCIM_AI\device_monitor\device_convert"
# ==============================================================

# 自动遍历所有 Excel 文件
for filename in os.listdir(FOLDER):
    # 只处理 .xlsx 文件
    if filename.lower().endswith(".xlsx"):
        print(f"\n正在处理：{filename}")

        try:
            # 拼接完整文件路径
            xlsx_file = os.path.join(FOLDER, filename)

            # 读取 Excel（唯一ID 等点分列禁止被推断为 float，否则 93.x.x.2.7.1.1 会失真）
            df = pd.read_excel(xlsx_file, engine="openpyxl")
            for col in df.columns:
                cn = unicodedata.normalize("NFKC", str(col).strip())
                if cn == "唯一ID" or ("唯一" in cn and "ID" in cn.upper()):
                    df[col] = df[col].map(
                        lambda x: (
                            ""
                            if pd.isna(x)
                            else unicodedata.normalize("NFKC", str(x).strip())
                        )
                    )

            # 生成同名、同路径的 .pkl 文件
            base_name = os.path.splitext(filename)[0]
            pkl_file = os.path.join(FOLDER, base_name + ".pkl")

            # 保存为高速格式
            df.to_pickle(pkl_file)
            print(f"✅ 转换成功 → {pkl_file}")

        except Exception as e:
            print(f"❌ 处理失败：{str(e)[:60]}")

print("\n🎉 全部转换完成！所有文件都保存在原文件夹，同名只改后缀 .pkl")