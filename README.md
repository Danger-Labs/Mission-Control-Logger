# Mission-Control-Logger
An automated Python dashboard for Omega wireless temperature probes. Features real-time live graphing and automated Excel export to eliminate manual data processing in high-volume environments.

## 📋 Overview
In fast-paced, high-volume laboratory and manufacturing environments, manual temperature logging and data post-processing create significant bottlenecks. **Mission Control** is a custom Python-based automation tool designed to eliminate these inefficiencies by interfacing directly with Omega UWTC wireless transmitters.

This system provides a live-view dashboard while silently exporting and graphing data directly into Excel, ensuring zero manual post-processing and protecting data integrity.

```mermaid
flowchart TD
    subgraph Phase1 [Phase 1: Data Transmission]
        A[Omega Wireless Logger] -- Raw °F Radio Signal --> B[USB Receiver COM4]
    end

    subgraph Phase2 [Phase 2: Live Processing]
        C{Mission Control Script}
        D[Identify Logger Address]
        E[Filter 9998 Error Spikes]
        F[Pure Math °F to °C Conversion]
        G[Write to CSV or RAM Buffer]
    end

    subgraph Phase3 [Phase 3: Live Dashboard]
        H[Assign Dashboard Color]
        I[Plot Live °C Data Point]
        J[Auto-Scale ±20% Y-Axis]
    end

    subgraph Phase4 [Phase 4: Post-Processing]
        K{Excel Auto-Generator}
        L[Read Completed Temp CSV]
        M[Generate Official .xlsx]
        N[Delete Temp CSV]
    end

    subgraph Phase5 [Phase 5: Final Storage]
        O[(Local Export Directory)]
        P[(Local Backup Directory)]
        Q[Log to Usage_History_Log.csv]
    end

    %% Connections
    B --> C
    C --> D
    D --> E
    E --> F
    F --> G
    
    G --> H
    H --> I
    I --> J
    
    G -. 5-Min Timeout .-> K
    K --> L
    L --> M
    M --> N
    
    M --> O
    M --> P
    M --> Q
```

## ✨ Key Features
* **Automated Data Pipeline:** Eliminates manual data handling by automatically exporting and formatting logger data directly into Excel spreadsheets.
* **Live-View Visualization:** Real-time monitoring of multiple Omega temperature probes simultaneously on a unified dashboard.
* **"Stealth Mode" Operation:** Engineered to run in the background (utilizing a `.pyw` architecture) to prevent accidental system closure by non-technical staff in active environments.
* **Zero-Downtime Hot-Reloading:** Allows for the seamless integration of new hardware and configuration updates without needing to restart the primary logging system.
* **End-User Focused:** Deployed with comprehensive checklists and Standard Operating Procedures (SOPs) for smooth onboarding and operational consistency.

## 🛠️ Tech Stack & Hardware
* **Language:** Python
* **Hardware Integration:** Omega UWTC Wireless Temperature Transmitters & Receivers
* **Output:** Automated Excel integration

## 🚀 Deployment & Documentation
*(Note: Upload your SOP and checklist files to the repo, then link them here)*
* [Standard Operating Procedure (SOP) - User Guide](#)
* [System Deployment Checklist](#)

## 🚧 Current Roadmap
* Fine-tuning stealth mode file paths and automated shortcut deployments.
* Refining the hot-reload logic for even faster hardware pairing.
* Optimizing the deployment process for new user onboarding.
