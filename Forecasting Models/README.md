# Time Series Forecasting & Modeling Suite

## Overview
This repository contains a comprehensive exploration of time-series forecasting and regression techniques. The project benchmarks various modeling approaches—ranging from classical statistical methods (ARIMA) to modern machine learning (CatBoost) and deep learning (RNN/LSTM)—and deploys them via interactive Gradio interfaces.

## System Architecture
The project follows a modular pipeline from raw data processing to model deployment.

```mermaid
graph TD
    A[Raw Data] --> B(EDA & Preprocessing)
    B --> C{Modeling Strategy}
    
    subgraph "Experimentation Phase"
    C --> D[Linear Models]
    C --> E[Time Series: ARIMA / Prophet]
    C --> F[Tree-based: CatBoost]
    C --> G[Deep Learning: RNN / LSTM]
    end
    
    D --> H[Model Artifacts]
    E --> H
    F --> H
    G --> H
    
    H --> I[Gradio Interfaces]
    I --> J((User Demo))