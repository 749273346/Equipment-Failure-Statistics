# 🍎 UI Beautification & Feature Upgrade Plan

I will modernize the application interface to mimic the clean, minimal aesthetic of macOS and introduce practical new features.

## 1. UI/UX Overhaul (Apple-Style Layout)
We will transition from the traditional "Tabbed" interface to a modern **Sidebar Navigation** layout.

*   **Sidebar Navigation**: A fixed left-hand panel containing navigation items (Data Collection, Statistics, Settings) with clean iconography and hover effects.
*   **Card-Based Design**: Content areas will use a "Card" metaphor (white background, subtle borders/shadows, rounded corners) set against a light gray application background (`#F5F5F7`).
*   **Typography**: Update fonts to use clean, modern sans-serif fonts (Microsoft YaHei UI / Segoe UI) with improved spacing and hierarchy.
*   **Color Palette**:
    *   **Background**: Light Gray (#F5F5F7)
    *   **Accent**: macOS Blue (#007AFF) for primary buttons and active states.
    *   **Status Colors**: Traffic light colors for status indicators (Red/Yellow/Green).

## 2. New Features
I will add the following features to enhance usability and analysis capabilities:

*   **📅 Date/Time Filtering**:
    *   Add a filter control in the "Statistics" view to analyze defects by **Year** and **Month**.
    *   Allow users to view "All Time" or specific periods.
*   **📤 Export Charts**:
    *   Add an "Export Chart" button to save the generated statistical charts as PNG/JPG images for reporting.
*   **🌗 Dark Mode Support**:
    *   Implement a toggle to switch between Light and Dark themes (a signature macOS feature).

## 3. Technical Implementation
*   **Library**: I will use `ttkbootstrap` (a modern wrapper for tkinter) if available, or manually style `tkinter.ttk` widgets to achieve the desired look without external heavy dependencies.
*   **Refactoring**:
    *   Modify `auto_fill_defects.py` to restructure the `App` class.
    *   Split the UI into `Sidebar`, `MainContent`, and `View` classes.

## Verification
*   Verify that the "Data Collection" process still works correctly with the new UI.
*   Test the new "Date Filter" against the existing Excel data.
*   Confirm the "Dark Mode" toggle correctly updates all UI elements.
