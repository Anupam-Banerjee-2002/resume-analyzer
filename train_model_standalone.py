"""
Standalone script: benchmark SVC vs RandomForest on data/resume_dataset.csv
and print a comparison report. (app.py trains its own model automatically on
startup using the same TF-IDF + Calibrated LinearSVC approach — this script is
for offline experimentation/benchmarking only.)

Usage:
    pip install pandas scikit-learn --break-system-packages
    python train_model_standalone.py
"""
import re
import pandas as pd
from sklearn.preprocessing import LabelEncoder
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.model_selection import train_test_split, cross_val_score, StratifiedKFold
from sklearn.svm import LinearSVC
from sklearn.calibration import CalibratedClassifierCV
from sklearn.ensemble import RandomForestClassifier
from sklearn.metrics import accuracy_score, classification_report

df = pd.read_csv("data/resume_dataset.csv")  # columns: text, label
print("Dataset shape:", df.shape)
print(df["label"].value_counts())

le = LabelEncoder()
df["label_enc"] = le.fit_transform(df["label"])

tfidf = TfidfVectorizer(stop_words="english", max_features=8000, ngram_range=(1, 2))
X = tfidf.fit_transform(df["text"])
y = df["label_enc"].values

X_train, X_test, y_train, y_test = train_test_split(X, y, test_size=0.2, random_state=42, stratify=y)

svc = CalibratedClassifierCV(LinearSVC(C=2.0, class_weight="balanced", max_iter=5000), cv=5)
svc.fit(X_train, y_train)
print("\n[SVC]")
print(classification_report(y_test, svc.predict(X_test), target_names=le.classes_))

rf = RandomForestClassifier(n_estimators=300, class_weight="balanced", random_state=42, n_jobs=-1)
rf.fit(X_train, y_train)
print("\n[RandomForest]")
print(classification_report(y_test, rf.predict(X_test), target_names=le.classes_))
