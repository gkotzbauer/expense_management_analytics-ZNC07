import { useState, useEffect, useRef } from "react"
import "./App.css"
import * as XLSX from "xlsx"
import Chart from 'chart.js/auto'

interface FinancialData {
  Metric_ID: string
  Metric_Name: string
  Category: string
  Responsibility: string
  Value_2024_Jan_July: number
  Value_2025_Jan_July: number
  Growth_Rate_Decimal: number
  Growth_Rate_Percentage: number
}

interface ExpenseData {
  Category: string
  "Jul 2025": number
  "Anchor vs Prior Avg ($)": number
  "Margin Risk Assessment": string
  "Expense Growth Alignment": string
  "Efficiency Alert": string
  "Marketing Spend Efficiency": string
}

function App() {
  const [financialData, setFinancialData] = useState<FinancialData[]>([])
  const [expenseData, setExpenseData] = useState<ExpenseData[]>([])
  const [error, setError] = useState<string | null>(null)
  const chartRef = useRef<HTMLCanvasElement | null>(null)
  const chartInstance = useRef<any>(null)

  useEffect(() => {
    const fetchData = async () => {
      try {
        // Fetch Financial Performance Data
        const financialResponse = await fetch("/ZNC07 Financial Performance Data.xlsx")
        if (!financialResponse.ok) throw new Error("Failed to fetch Financial Performance Data")
        const financialBlob = await financialResponse.blob()
        const financialArrayBuffer = await financialBlob.arrayBuffer()
        const financialWorkbook = XLSX.read(financialArrayBuffer, { type: "array" })
        const financialSheetName = financialWorkbook.SheetNames[0]
        const financialWorksheet = financialWorkbook.Sheets[financialSheetName]
        const financialJsonData: FinancialData[] = XLSX.utils.sheet_to_json(financialWorksheet, { defval: "" })
        setFinancialData(financialJsonData)

        // Fetch Expense Analysis Data
        const expenseResponse = await fetch("/expense-analysis.xlsx")
        if (!expenseResponse.ok) throw new Error("Failed to fetch Expense Analysis Data")
        const expenseBlob = await expenseResponse.blob()
        const expenseArrayBuffer = await expenseBlob.arrayBuffer()
        const expenseWorkbook = XLSX.read(expenseArrayBuffer, { type: "array" })
        const expenseSheetName = expenseWorkbook.SheetNames[0]
        const expenseWorksheet = expenseWorkbook.Sheets[expenseSheetName]
        const expenseJsonData: ExpenseData[] = XLSX.utils.sheet_to_json(expenseWorksheet, { defval: "" })
        setExpenseData(expenseJsonData)
      } catch (err: any) {
        setError(err.message || "Error loading data")
      }
    }
    fetchData()
  }, [])

  useEffect(() => {
    if (!chartRef.current || expenseData.length === 0) return
    // Destroy previous chart instance if exists
    if (chartInstance.current) {
      chartInstance.current.destroy()
    }
    // Count occurrences of each value in 'Margin Risk Assessment'
    const counts: Record<string, number> = {}
    expenseData.forEach(row => {
      const key = row["Margin Risk Assessment"] || "Unknown"
      counts[key] = (counts[key] || 0) + 1
    })
    chartInstance.current = new Chart(chartRef.current, {
      type: 'bar',
      data: {
        labels: Object.keys(counts),
        datasets: [{
          label: 'Count',
          data: Object.values(counts),
          backgroundColor: ['#e53e3e', '#fd7900', '#48bb78', '#4299e1'],
        }]
      },
      options: {
        responsive: true,
        plugins: { legend: { display: false } },
        scales: { y: { beginAtZero: true, ticks: { stepSize: 1 } } }
      }
    })
  }, [expenseData])

  // Filter data by category
  const yoyData = financialData.filter(row => row.Category === "YOY Expense & Profitability Analysis")
  const cashflowData = financialData.filter(row => row.Category === "Remaining Year Cashflow Projections")

  // Format currency
  const formatCurrency = (value: number) => {
    if (typeof value !== 'number' || isNaN(value)) return '$0'
    return new Intl.NumberFormat('en-US', { style: 'currency', currency: 'USD' }).format(value)
  }

  // Format percentage
  const formatPercentage = (value: number) => {
    if (typeof value !== 'number' || isNaN(value)) return '0%'
    return `${(value * 100).toFixed(1)}%`
  }

  return (
    <div>
      <h1>ZNC07 Margin Performance Dashboard</h1>
      {error && <div style={{ color: "red" }}>{error}</div>}
      
      {/* YOY Expense & Profitability Analysis Table */}
      <div style={{ margin: "20px 0" }}>
        <h2>📊 YOY Expense & Profitability Analysis</h2>
        <table style={{ margin: "0 auto", borderCollapse: "collapse", width: "100%", maxWidth: "1200px" }}>
          <thead>
            <tr>
              <th style={{ border: "1px solid #ccc", padding: "8px", backgroundColor: "#f5f5f5" }}>Metric</th>
              <th style={{ border: "1px solid #ccc", padding: "8px", backgroundColor: "#f5f5f5" }}>Category</th>
              <th style={{ border: "1px solid #ccc", padding: "8px", backgroundColor: "#f5f5f5" }}>Responsibility</th>
              <th style={{ border: "1px solid #ccc", padding: "8px", backgroundColor: "#f5f5f5" }}>2024 (Jan-Jul)</th>
              <th style={{ border: "1px solid #ccc", padding: "8px", backgroundColor: "#f5f5f5" }}>2025 (Jan-Jul)</th>
              <th style={{ border: "1px solid #ccc", padding: "8px", backgroundColor: "#f5f5f5" }}>Growth Rate</th>
            </tr>
          </thead>
          <tbody>
            {yoyData.map((row, i) => (
              <tr key={i}>
                <td style={{ border: "1px solid #ccc", padding: "8px" }}>{row.Metric_Name}</td>
                <td style={{ border: "1px solid #ccc", padding: "8px" }}>{row.Category}</td>
                <td style={{ border: "1px solid #ccc", padding: "8px" }}>{row.Responsibility}</td>
                <td style={{ border: "1px solid #ccc", padding: "8px", textAlign: "right" }}>{formatCurrency(row.Value_2024_Jan_July)}</td>
                <td style={{ border: "1px solid #ccc", padding: "8px", textAlign: "right" }}>{formatCurrency(row.Value_2025_Jan_July)}</td>
                <td style={{ border: "1px solid #ccc", padding: "8px", textAlign: "right" }}>{formatPercentage(row.Growth_Rate_Decimal)}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>

      {/* Remaining Year Cashflow Projections Table */}
      <div style={{ margin: "20px 0" }}>
        <h2>💰 Remaining Year Cashflow Projections</h2>
        <table style={{ margin: "0 auto", borderCollapse: "collapse", width: "100%", maxWidth: "1200px" }}>
          <thead>
            <tr>
              <th style={{ border: "1px solid #ccc", padding: "8px", backgroundColor: "#f5f5f5" }}>Metric</th>
              <th style={{ border: "1px solid #ccc", padding: "8px", backgroundColor: "#f5f5f5" }}>Category</th>
              <th style={{ border: "1px solid #ccc", padding: "8px", backgroundColor: "#f5f5f5" }}>Responsibility</th>
              <th style={{ border: "1px solid #ccc", padding: "8px", backgroundColor: "#f5f5f5" }}>2024 (Jan-Jul)</th>
              <th style={{ border: "1px solid #ccc", padding: "8px", backgroundColor: "#f5f5f5" }}>2025 (Jan-Jul)</th>
              <th style={{ border: "1px solid #ccc", padding: "8px", backgroundColor: "#f5f5f5" }}>Growth Rate</th>
            </tr>
          </thead>
          <tbody>
            {cashflowData.map((row, i) => (
              <tr key={i}>
                <td style={{ border: "1px solid #ccc", padding: "8px" }}>{row.Metric_Name}</td>
                <td style={{ border: "1px solid #ccc", padding: "8px" }}>{row.Category}</td>
                <td style={{ border: "1px solid #ccc", padding: "8px" }}>{row.Responsibility}</td>
                <td style={{ border: "1px solid #ccc", padding: "8px", textAlign: "right" }}>{formatCurrency(row.Value_2024_Jan_July)}</td>
                <td style={{ border: "1px solid #ccc", padding: "8px", textAlign: "right" }}>{formatCurrency(row.Value_2025_Jan_July)}</td>
                <td style={{ border: "1px solid #ccc", padding: "8px", textAlign: "right" }}>{formatPercentage(row.Growth_Rate_Decimal)}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>

      {/* Most Recent Month Key Insights */}
      <div style={{ margin: "20px 0" }}>
        <h2>📈 Most Recent Month Key Insights (July 2025)</h2>
        <table style={{ margin: "0 auto", borderCollapse: "collapse", width: "100%", maxWidth: "1200px" }}>
          <thead>
            <tr>
              <th style={{ border: "1px solid #ccc", padding: "8px", backgroundColor: "#f5f5f5" }}>Category</th>
              <th style={{ border: "1px solid #ccc", padding: "8px", backgroundColor: "#f5f5f5" }}>Expenses for Most Recent Month</th>
              <th style={{ border: "1px solid #ccc", padding: "8px", backgroundColor: "#f5f5f5" }}>Most Recent Month vs. Avg. of Prior Months</th>
              <th style={{ border: "1px solid #ccc", padding: "8px", backgroundColor: "#f5f5f5" }}>Margin Risk Assessment</th>
              <th style={{ border: "1px solid #ccc", padding: "8px", backgroundColor: "#f5f5f5" }}>Expense Growth Alignment</th>
              <th style={{ border: "1px solid #ccc", padding: "8px", backgroundColor: "#f5f5f5" }}>Suggested Action</th>
            </tr>
          </thead>
          <tbody>
            {(() => {
              // Order for Suggested Action
              const actionOrder = [
                "Investigate – Potential Risk",
                "Efficient Scaling",
                "Productivity Gain",
                "Strong Cost Control",
                "Review Volatility",
                "Stable – No Action",
                "No Comparison",
                "Below Threshold"
              ];
              // Sort data by action order
              const sorted = [...expenseData].sort((a, b) => {
                const aIdx = actionOrder.indexOf(a["Efficiency Alert"] || "")
                const bIdx = actionOrder.indexOf(b["Efficiency Alert"] || "")
                return (aIdx === -1 ? 999 : aIdx) - (bIdx === -1 ? 999 : bIdx)
              });
              return sorted.map((row, i) => (
                <tr key={i}>
                  <td style={{ border: "1px solid #ccc", padding: "8px" }}>{row["Category"]}</td>
                  <td style={{ border: "1px solid #ccc", padding: "8px", textAlign: "right" }}>{formatCurrency(row["Jul 2025"])}</td>
                  <td style={{ border: "1px solid #ccc", padding: "8px", textAlign: "right" }}>{formatCurrency(row["Anchor vs Prior Avg ($)"])}</td>
                  <td style={{ border: "1px solid #ccc", padding: "8px" }}>{row["Margin Risk Assessment"]}</td>
                  <td style={{ border: "1px solid #ccc", padding: "8px" }}>{row["Expense Growth Alignment"]}</td>
                  <td style={{ border: "1px solid #ccc", padding: "8px" }}>{row["Efficiency Alert"]}</td>
                </tr>
              ))
            })()}
          </tbody>
        </table>
      </div>

      {/* Key Insights Section */}
      <div style={{ margin: "20px 0" }}>
        <h2>🔍 Key Insights</h2>
        <div style={{ margin: "20px 0" }}>
          <h3>Marketing Spend Efficiency</h3>
          <p><strong>Advertising & Marketing Fund (Inefficient, 0.06)</strong></p>
        </div>
        
        <div style={{ margin: "20px 0" }}>
          <h3>Note</h3>
          <p>Note: The value behind the category is the ratio between the percentage growth for the category and the percentage growth for Gross Profit. A negative number means the category grew slower than Gross Profit. A positive number means the growth in the expense outpaced the growth in revenue (gross profit). For example, a value of 0.06 means the category grew at 6% of the rate the revenue growth.</p>
        </div>
      </div>

      {/* Chart Section */}
      <div style={{ margin: "20px 0" }}>
        <h2>📊 Margin Risk Assessment Chart</h2>
        <div style={{ maxWidth: 600, margin: '0 auto' }}>
          <canvas ref={chartRef} height={300} />
        </div>
      </div>
    </div>
  )
}

export default App
