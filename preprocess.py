import pandas as pd
import numpy as np
from sklearn.model_selection import train_test_split
from sklearn.preprocessing import StandardScaler
import joblib
import os

# Columns where 0 is biologically impossible and means "missing"
ZERO_MEANS_MISSING = ['Glucose', 'BloodPressure', 'SkinThickness', 'Insulin', 'BMI']

def load_data(filepath):
    df = pd.read_csv(filepath)
    return df

def replace_zeros_with_median(df):
    df = df.copy()
    for col in ZERO_MEANS_MISSING:
        median = df[col].replace(0, np.nan).median()
        df[col] = df[col].replace(0, median)
    return df

def split_data(df):
    X = df.drop('Outcome', axis=1)
    y = df['Outcome']

    # First split: 70% train, 30% temp
    X_train, X_temp, y_train, y_temp = train_test_split(
        X, y, test_size=0.30, random_state=42, stratify=y
    )
    # Second split: split temp equally into 15% val and 15% test
    X_val, X_test, y_val, y_test = train_test_split(
        X_temp, y_temp, test_size=0.50, random_state=42, stratify=y_temp
    )
    return X_train, X_val, X_test, y_train, y_val, y_test

def scale_features(X_train, X_val, X_test, scaler_path=None):
    scaler = StandardScaler()
    X_train_scaled = scaler.fit_transform(X_train)
    X_val_scaled = scaler.transform(X_val)
    X_test_scaled = scaler.transform(X_test)

    if scaler_path:
        os.makedirs(os.path.dirname(scaler_path), exist_ok=True)
        joblib.dump(scaler, scaler_path)

    return X_train_scaled, X_val_scaled, X_test_scaled, scaler

def preprocess_pipeline(filepath, scaler_path=None):
    df = load_data(filepath)
    df = replace_zeros_with_median(df)
    X_train, X_val, X_test, y_train, y_val, y_test = split_data(df)
    X_train_s, X_val_s, X_test_s, scaler = scale_features(
        X_train, X_val, X_test, scaler_path
    )
    return X_train_s, X_val_s, X_test_s, y_train, y_val, y_test, scaler
