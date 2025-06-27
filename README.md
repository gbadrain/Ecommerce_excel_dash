# Ecommerce Sales Analysis Dashboard 

An interactive Excel-based dashboard that provides comprehensive insights into ecommerce sales performance across regions, time periods, and product categories. Built with Excel PivotTables, charts, and VBA automation for dynamic data exploration.
<img width="1621" alt="Screenshot 2025-06-26 at 3 25 50 PM" src="https://github.com/user-attachments/assets/1982416c-7c39-4712-bf97-5c9b4650a065" />

## Features

### Key Performance Indicators
- **Sales Metrics**: Total sales, profit, quantity sold, and order count
- **Profit Margin Analysis**: Comprehensive profitability insights
- **Year-over-Year Growth**: Performance tracking from 2011-2014

### Interactive Visualizations
- **Monthly Trends**: Dual-axis combo charts showing sales and profit patterns
- **Category Analysis**: Waterfall charts for profit breakdown and pie charts for sales distribution
- **Product Performance**: Top 5 sub-categories ranking by sales volume
- **Geographic Insights**: US state-level sales mapping
- **Dynamic Filtering**: Timeline and slicer controls for real-time data exploration

### Automation Features
- **One-Click Filter Reset**: VBA macro for instant filter clearing
- **Presentation Mode**: Full-screen visualization with hidden ribbon and gridlines
- **Interactive Controls**: User-friendly slicers and timeline filters

## Project Structure

```
Ecommerce_Dashboard/
├── Ecommerce Sales Analysis.xlsm         # Main interactive dashboard with macros
└── Worksheet images/                      # Dashboard screenshots
    ├── Screenshot 2025-06-26 at 3.25.50 PM.png
    ├── Screenshot 2025-06-26 at 3.26.58 PM.png
    ├── Screenshot 2025-06-26 at 3.27.05 PM.png
    ├── Screenshot 2025-06-26 at 3.27.10 PM.png
    ├── Screenshot 2025-06-26 at 3.27.18 PM.png
    ├── Screenshot 2025-06-26 at 3.27.24 PM.png
    ├── Screenshot 2025-06-26 at 3.27.48 PM.png
    └── Screenshot 2025-06-26 at 3.27.55 PM.png
```

## Technologies Used

- **Platform**: Microsoft Excel (.xlsm format)
- **Automation**: VBA (Visual Basic for Applications)
- **Data Modeling**: PivotTables and Excel Tables
- **Visualizations**: PivotCharts (Column, Line, Pie, Waterfall, Geographic Maps)
- **Interactivity**: Timeline controls and slicers
- **Business Intelligence**: KPI cards and performance metrics

## Dashboard Components

### KPI Summary Cards
- Total Sales Revenue
- Profit Analysis
- Quantity Metrics
- Order Count
- Profit Margin Percentage

### Analytical Views
- **Monthly Sales & Profit Trends**: Track performance over time
- **Category Performance**: Waterfall chart showing profit contribution by category
- **Sales Distribution**: Pie chart visualization of sales share
- **Top Products**: Ranking of best-performing sub-categories
- **Regional Analysis**: Geographic distribution of sales across US states

## Getting Started

### Prerequisites
- Microsoft Excel (Desktop version recommended)
- Windows or Mac with Excel macro support enabled

### Installation & Usage

1. **Download the Dashboard**
   ```
   Download: Ecommerce Sales Analysis.xlsm
   ```

2. **Enable Macros**
   - Open the file in Excel
   - Click "Enable Content" when prompted to activate VBA macros

3. **Explore the Data**
   - Use slicers to filter by year and region
   - Interact with the timeline to focus on specific date ranges
   - Analyze different views using the various chart types

4. **Reset Filters**
   - Click the "Reset Filters" button to clear all active filters
   - Instantly return to the full dataset view

5. **Presentation Mode**
   - Press `Ctrl + F1` to hide the Excel ribbon
   - Enter full-screen mode for professional presentations

## VBA Code Reference

### Reset Filters Macro
```vb
Sub ResetAllFilters()
    Const xlTimeline As Long = 7
    Dim sc As SlicerCache
    For Each sc In ThisWorkbook.SlicerCaches
        If sc.SlicerCacheType = xlTimeline Then
            sc.TimelineState.ClearAllFilters
        Else
            sc.ClearManualFilter
        End If
    Next sc
End Sub
```

This macro efficiently clears both timeline and slicer filters with a single button click.

## Use Cases

- **Business Intelligence**: Executive reporting and KPI monitoring
- **Sales Analysis**: Performance tracking across regions and time periods
- **Product Management**: Category and sub-category performance evaluation
- **Strategic Planning**: Year-over-year growth analysis and trend identification
- **Presentation**: Client-ready dashboards for business reviews

## Contributing

This project demonstrates Excel and VBA capabilities for business intelligence. Contributions, suggestions, and feedback are welcome!

## Contact

**Gurpreet Singh Badrain**  
*Market Research Analyst & Aspiring Data Analyst*

- **Portfolio**: [Data Guru](https://datascienceportfol.io/gbadrain)
- **LinkedIn**: [gurpreet-badrain](http://linkedin.com/in/gurpreet-badrain-b258a0219)
- **Email**: gbadrain@gmail.com
- **GitHub**: [gbadrain](https://github.com/gbadrain)
- **Streamlit**: [gbadrain-Machine Learning](https://gbadrain-machine-learning.streamlit.app)

## Acknowledgments

- **Copilot AI** for development assistance
- **YouTube Channel**: [datatutorials1](https://www.youtube.com/@datatutorials1) for inspiration and learning

---
⭐ Star this repository if you find it helpful!**
