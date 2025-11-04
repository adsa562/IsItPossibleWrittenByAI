import os
import numpy as np
import pandas as pd
import textstat
import nltk
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
import subprocess
from nltk import word_tokenize, sent_tokenize
from collections import Counter
from docx import Document
from tkinter import Tk, filedialog, messagebox
from tqdm import tqdm

nltk.download('punkt', quiet=True)

# ==============================
# 設定字型（微軟正黑體）
# ==============================
def set_chinese_font():
    font_list = [f.name for f in fm.fontManager.ttflist]
    if "Microsoft JhengHei" in font_list:
        plt.rcParams["font.sans-serif"] = ["Microsoft JhengHei"]
    elif "SimHei" in font_list:
        plt.rcParams["font.sans-serif"] = ["SimHei"]  # 備用黑體
    else:
        plt.rcParams["font.sans-serif"] = ["Arial Unicode MS"]
    plt.rcParams["axes.unicode_minus"] = False  # 避免負號變成方塊

set_chinese_font()

# ==============================
# AI 文本特徵分析核心
# ==============================
def analyze_text_features(text):
    if not text or len(text.strip()) < 50:
        return None

    sentences = sent_tokenize(text)
    words = word_tokenize(text.lower())
    if not words:
        return None

    # 詞彙多樣性（越低代表越像AI）
    ttr = len(set(words)) / len(words)
    # 句長分布與波動（越平均代表越像AI）
    sent_lengths = [len(word_tokenize(s)) for s in sentences]
    burstiness = np.std(sent_lengths) / np.mean(sent_lengths)
    # 可讀性（太高或太低都可能是AI）
    readability = textstat.flesch_reading_ease(text)
    # 重複度（AI常重複詞）
    common_ratio = Counter(words).most_common(1)[0][1] / len(words)
    # 粗略困惑度（以句長變化代替）
    pseudo_perplexity = np.var(sent_lengths)

    score = (
        (1 - ttr) * 0.3 +
        (1 - burstiness) * 0.2 +
        common_ratio * 0.2 +
        (pseudo_perplexity < 100) * 0.3
    )
    ai_score = round(score * 100, 2)

    if ai_score < 40:
        result = "人類撰寫"
    elif ai_score < 70:
        result = "模糊區（混合或修飾過）"
    else:
        result = "高機率 AI 生成"

    return {
        "文字長度": len(text),
        "句子數": len(sentences),
        "詞彙多樣性": round(ttr, 3),
        "句長變異": round(burstiness, 3),
        "可讀性": round(readability, 2),
        "重複率": round(common_ratio, 3),
        "困惑度代理": round(pseudo_perplexity, 2),
        "AI可能性分數": ai_score,
        "分析結果": result
    }

# ==============================
# 讀取文件內容
# ==============================
def read_file_content(filepath):
    if filepath.endswith(".txt"):
        with open(filepath, "r", encoding="utf-8", errors="ignore") as f:
            return f.read()
    elif filepath.endswith(".docx"):
        doc = Document(filepath)
        return "\n".join([p.text for p in doc.paragraphs])
    return ""

# ==============================
# 主分析流程
# ==============================
def analyze_folder(folder_path, output_path):
    results = []
    files = []
    for root, _, filelist in os.walk(folder_path):
        for f in filelist:
            if f.endswith((".txt", ".docx")):
                files.append(os.path.join(root, f))

    if not files:
        messagebox.showwarning("提示", "找不到 .txt 或 .docx 檔案")
        return

    for file in tqdm(files, desc="分析中文章中..."):
        text = read_file_content(file)
        analysis = analyze_text_features(text)
        if analysis:
            analysis["檔案名稱"] = os.path.basename(file)
            results.append(analysis)

    if not results:
        messagebox.showinfo("結果", "沒有可分析的內容。")
        return

    df = pd.DataFrame(results)
    df = df[["檔案名稱", "文字長度", "句子數", "詞彙多樣性", "句長變異",
             "可讀性", "重複率", "困惑度代理", "AI可能性分數", "分析結果"]]

    # ======================
    # 統計摘要
    # ======================
    summary = {
        "分析文件總數": len(df),
        "平均 AI 分數": round(df["AI可能性分數"].mean(), 2),
        "最高分": df["AI可能性分數"].max(),
        "最低分": df["AI可能性分數"].min(),
        "高機率 AI 數": sum(df["分析結果"] == "高機率 AI 生成"),
        "模糊區數": sum(df["分析結果"].str.contains("模糊")),
        "明顯人類撰寫數": sum(df["分析結果"] == "人類撰寫")
    }
    summary_df = pd.DataFrame([summary])

    # ======================
    # 寫入 Excel
    # ======================
    output_excel = os.path.join(output_path, "AI文本分析報告.xlsx")
    with pd.ExcelWriter(output_excel, engine="openpyxl") as writer:
        summary_df.to_excel(writer, sheet_name="摘要", index=False)
        df.to_excel(writer, sheet_name="詳細結果", index=False)

    # ======================
    # 產生統計圖表（使用微軟正黑體）
    # ======================
    plt.figure(figsize=(10, 6))
    plt.barh(df["檔案名稱"], df["AI可能性分數"], color="#4682B4")
    plt.xlabel("AI 生成可能性分數", fontsize=12)
    plt.ylabel("檔案名稱", fontsize=12)
    plt.title("AI 文本偵測分析結果", fontsize=14, fontweight="bold")

    avg_score = df["AI可能性分數"].mean()
    plt.axvline(avg_score, color="red", linestyle="--", label=f"平均值 {avg_score:.2f}")
    plt.legend()
    plt.tight_layout()

    output_chart = os.path.join(output_path, "AI_score_chart.png")
    plt.savefig(output_chart, dpi=200)
    plt.close()

    # ======================
    # 自動開啟輸出資料夾
    # ======================
    try:
        if os.name == "nt":  # Windows
            os.startfile(output_path)
        elif os.name == "posix":  # macOS / Linux
            subprocess.Popen(["xdg-open", output_path])
    except Exception as e:
        print(f"⚠️ 無法自動開啟資料夾：{e}")

    messagebox.showinfo(
        "完成",
        f"✅ 分析完成！\n\n報告輸出：{output_excel}\n圖表輸出：{output_chart}\n\n已自動開啟輸出資料夾。\n📊注意：本工具僅供參考，請勿作為唯一判斷AI文筆與否的依據。📊 \n\n輸出解讀參考：\n0–40％ → 很可能是人類撰寫\n40–70％ → 模糊區（可能混合）\n70–100％ → 高機率為 AI 生成"
    )

# ==============================
# GUI 主介面
# ==============================
def main_gui():
    root = Tk()
    root.withdraw()
    messagebox.showinfo(
        "AI 文本自動檢測器 v1.0",
        "此工具可分析資料夾內的 .txt 與 .docx 文件，\n判斷是否可能由 AI 生成，並輸出統計報告與圖表。\nGitHub: https://github.com/adsa562/IsItPossibleWrittenByAI"
    )

    folder_path = filedialog.askdirectory(title="請選擇要分析的資料夾")
    if not folder_path:
        messagebox.showinfo("提示", "未選擇分析資料夾，程式結束。")
        root.destroy()
        return

    output_path = filedialog.askdirectory(title="請選擇報告輸出路徑")
    if not output_path:
        messagebox.showinfo("提示", "未選擇輸出路徑，程式結束。")
        root.destroy()
        return

    analyze_folder(folder_path, output_path)
    root.destroy()

if __name__ == "__main__":
    main_gui()
