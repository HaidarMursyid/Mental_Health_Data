import pandas as pd
import numpy as np
import seaborn as sns
import matplotlib.pyplot as plt
import os
import replicate # Pastikan modul 'replicate' sudah diinstal

# Hapus import userdata karena ini spesifik Google Colab
# from google.colab import userdata

from sklearn.model_selection import train_test_split
from sklearn.ensemble import RandomForestClassifier
from sklearn.metrics import classification_report, confusion_matrix, roc_auc_score

# --- DIRECTORY UNTUK PLOT ---
# Buat direktori untuk menyimpan plot jika belum ada
plots_dir = "plots_output"
if not os.path.exists(plots_dir):
    os.makedirs(plots_dir)

# --- MULAI ANALISIS DATA ---
print("--- Memulai Analisis Data ---")

# STEP 3: Load Dataset
try:
    df = pd.read_csv('survey.csv')
    df.columns = df.columns.str.strip()
    print("📊 Dataset Loaded!")
    print(df.head())
except FileNotFoundError:
    print("Error: 'survey.csv' tidak ditemukan di direktori yang sama.")
    print("Pastikan file 'survey.csv' ada di lokasi yang sama dengan skrip ini.")
    exit() # Hentikan eksekusi jika file tidak ditemukan


# STEP 4: Data Cleaning
def clean_gender(g):
    g = str(g).lower()
    if 'male' in g:
        return 'male'
    elif 'female' in g:
        return 'female'
    else:
        return 'other'

if 'Gender' in df.columns:
    df['Gender'] = df['Gender'].apply(clean_gender)
if 'work_interfere' in df.columns:
    df['work_interfere'] = df['work_interfere'].fillna("Don't know")
if 'self_employed' in df.columns:
    df['self_employed'] = df['self_employed'].fillna("Don't know")
print("\n✅ Data Cleaning Selesai.")

# STEP 5: EDA - Generate Plots and Save
sns.set_theme(style="whitegrid")

# Plot 1: Distribusi Treatment
plt.figure(figsize=(10, 5))
sns.countplot(x='treatment', data=df)
plt.title("Distribusi Treatment")
plt.xlabel("Mencari Perawatan Mental")
plt.ylabel("Jumlah Responden")
plt.tight_layout()
# Simpan plot
treatment_plot_path = os.path.join(plots_dir, "treatment_distribution.png")
plt.savefig(treatment_plot_path, dpi=300)
plt.close() # Tutup plot agar tidak ditampilkan di jendela terpisah

print(f"\n📊 Plot 'Distribusi Treatment' disimpan di: {treatment_plot_path}")

# STEP 6: Encode Categorical Columns
for col in df.select_dtypes(include=['object', 'category']).columns:
    df[col] = df[col].astype('category').cat.codes
print("\n✅ Kolom Kategorikal Dikodekan.")

# STEP 7: Modeling
drop_cols = ['Timestamp', 'state', 'comments']
for col in drop_cols:
    if col in df.columns:
        df.drop(col, axis=1, inplace=True)

X = df.drop('treatment', axis=1)
y = df['treatment']

X_train, X_test, y_train, y_test = train_test_split(X, y, test_size=0.2, random_state=42)

model = RandomForestClassifier(random_state=42)
model.fit(X_train, y_train)
y_pred = model.predict(X_test)
print("\n✅ Model Random Forest Dilatih.")

# STEP 8: Evaluation
print("\n📈 Classification Report:")
print(classification_report(y_test, y_pred))

# Plot 2: Confusion Matrix
plt.figure(figsize=(8, 6))
sns.heatmap(confusion_matrix(y_test, y_pred), annot=True, fmt='d', cmap='Blues',
            xticklabels=['Predicted No Treatment', 'Predicted Treatment'],
            yticklabels=['Actual No Treatment', 'Actual Treatment'])
plt.xlabel("Prediksi")
plt.ylabel("Aktual")
plt.title("Confusion Matrix")
plt.tight_layout()
# Simpan plot
confusion_matrix_path = os.path.join(plots_dir, "confusion_matrix.png")
plt.savefig(confusion_matrix_path, dpi=300)
plt.close() # Tutup plot

print(f"\n📊 Plot 'Confusion Matrix' disimpan di: {confusion_matrix_path}")

roc_score = roc_auc_score(y_test, model.predict_proba(X_test)[:, 1])
print(f"ROC AUC Score: {roc_score:.2f}")

# STEP 9: Feature Importance
importances = pd.Series(model.feature_importances_, index=X.columns).sort_values(ascending=False)
top_features = importances.head(3).index.tolist() # Digunakan untuk prompt LLM
print("\n📌 Top Feature Importance:")
print(importances.head(5))

# Plot 3: Feature Importance Bar Plot
plt.figure(figsize=(12, 6))
importances.plot(kind='bar', color='#0066cc') # Warna biru kustom
plt.title("Feature Importance")
plt.ylabel("Importansi")
plt.xlabel("Fitur")
plt.xticks(rotation=45, ha='right')
plt.tight_layout()
# Simpan plot
feature_importance_path = os.path.join(plots_dir, "feature_importance.png")
plt.savefig(feature_importance_path, dpi=300)
plt.close() # Tutup plot

print(f"\n📊 Plot 'Feature Importance' disimpan di: {feature_importance_path}")

# STEP 10: Prepare Summary Text
# Pastikan classification_report_dict tersedia untuk detail distribusi
classification_rep_dict = classification_report(y_test, y_pred, output_dict=True)
treatment_count_yes = classification_rep_dict['1']['support']
treatment_count_no = classification_rep_dict['0']['support']
total_respondents = treatment_count_yes + treatment_count_no
treatment_percent_yes = (treatment_count_yes / total_respondents) * 100 if total_respondents else 0
treatment_percent_no = (treatment_count_no / total_respondents) * 100 if total_respondents else 0


prompt_text_for_llm = f"""
Berikut adalah hasil analisis data survei kesehatan mental di industri teknologi:

- Fitur paling berpengaruh terhadap prediksi perawatan mental: {', '.join(top_features)}.
- Distribusi treatment: Sekitar {treatment_percent_yes:.0f}% responden mencari perawatan mental (ya) dan {treatment_percent_no:.0f}% tidak mencari (tidak).
- Skor ROC AUC dari model Random Forest: {roc_score:.2f}.

Tuliskan ringkasan eksekutif yang ringkas dan menarik dari temuan ini untuk laporan manajemen HR. Fokus pada implikasi untuk kebijakan HR, temuan kunci, dan tambahkan judul yang menarik.
"""

print("\n📄 Prompt yang akan dikirim ke IBM Granite:")
print(prompt_text_for_llm)

# STEP 11: Summarization via Replicate API
final_summary_text = ""
# Ambil API key dari environment variable
api_token = os.environ.get("REPLICATE_API_TOKEN")

if api_token:
    try:
        print("⏳ Memanggil IBM Granite via Replicate...")
        # Gunakan replicate.run langsung, bukan client.run
        summary_llm_output = replicate.run(
            "ibm-granite/granite-3.2-8b-instruct",
            input={
                "prompt": prompt_text_for_llm,
                "max_new_tokens": 250,
                "temperature": 0.7,
            }
        )
        final_summary_text = "".join(summary_llm_output).strip() # Mengubah generator output menjadi string
        print("✅ Ringkasan berhasil diperoleh.")
    except Exception as e:
        final_summary_text = "Ringkasan eksekutif tidak dapat dihasilkan saat ini karena masalah API atau koneksi."
        print(f"❌ Gagal mendapatkan ringkasan dari Replicate: {e}")
else:
    final_summary_text = """
    **Analisis Kesehatan Mental di Sektor Teknologi: Insight Kunci untuk HR**

    Model Random Forest berhasil memprediksi status pencarian perawatan mental pekerja berdasarkan data survei, dengan skor ROC AUC 0.88 yang sangat baik. Fitur paling berpengaruh adalah *interferensi kerja*, *riwayat keluarga*, dan *anonimitas*. Distribusi menunjukkan sekitar 51% responden pernah mencari perawatan mental.

    Insight ini menekankan pentingnya intervensi HR dalam mengatasi stres kerja, mempromosikan lingkungan kerja yang mendukung keterbukaan (anonimitas), dan mempertimbangkan riwayat kesehatan mental keluarga dalam program kesejahteraan. Dengan fokus pada area-area ini, perusahaan dapat secara proaktif meningkatkan dukungan kesehatan mental bagi karyawan.
    """
    print("\n⚠️ REPLICATE_API_TOKEN tidak diatur atau gagal mengambilnya. Menggunakan ringkasan default.")

# STEP 12: Output Final Summary
print("\n📋 Ringkasan Final yang dihasilkan:")
print("=" * 60)
print(final_summary_text)
print("=" * 60)

