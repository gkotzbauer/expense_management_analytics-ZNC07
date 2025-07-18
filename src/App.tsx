import { useState, useEffect, useRef } from "react"
import "./App.css"
import * as XLSX from "xlsx"
import Chart from 'chart.js/auto'

function App() {
  const [data, setData] = useState<any[]>([])
  const [error, setError] = useState<string | null>(null)
  const chartRef = useRef<HTMLCanvasElement | null>(null)
  const chartInstance = useRef<any>(null)

  useEffect(() => {
    const fetchExcel = async () => {
      try {
        const response = await fetch("/expense-analysis.xlsx")
        if (!response.ok) throw new Error("Failed to fetch Excel file")
        const blob = await response.blob()
        const arrayBuffer = await blob.arrayBuffer()
        const workbook = XLSX.read(arrayBuffer, { type: "array" })
        const sheetName = workbook.SheetNames[0]
        const worksheet = workbook.Sheets[sheetName]
        const jsonData: any[] = XLSX.utils.sheet_to_json(worksheet, { defval: "" })
        // Map 'Margin Risk Assessment' to 'Performance Diagnostic Summary'
        const mappedData = jsonData.map(row => ({
          ...row,
          "Performance Diagnostic Summary": row["Margin Risk Assessment"] || "Unknown",
          "Elasticity Classification": String(row["Elasticity Classification"] || "").trim()
        }))
        setData(mappedData)
      } catch (err: any) {
        setError(err.message || "Error loading data")
      }
    }
    fetchExcel()
  }, [])

  useEffect(() => {
    if (!chartRef.current || data.length === 0) return
    // Destroy previous chart instance if exists
    if (chartInstance.current) {
      chartInstance.current.destroy()
    }
    // Count occurrences of each value in 'Performance Diagnostic Summary'
    const counts: Record<string, number> = {}
    data.forEach(row => {
      const key = row["Performance Diagnostic Summary"] || "Unknown"
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
  }, [data])

  return (
    <div>
      <h1>High Point Margin Performance Dashboard</h1>
      {error && <div style={{ color: "red" }}>{error}</div>}
      <div style={{ maxWidth: 600, margin: '0 auto' }}>
        <h2>📊 Margin Risk Assessment</h2>
        <canvas ref={chartRef} height={300} />
      </div>
      <table style={{ margin: "0 auto", borderCollapse: "collapse" }}>
        <thead>
          <tr>
            <th style={{ border: "1px solid #ccc", padding: "4px 8px" }}>Category</th>
            <th style={{ border: "1px solid #ccc", padding: "4px 8px" }}>Expenses for Most Recent Month</th>
            <th style={{ border: "1px solid #ccc", padding: "4px 8px" }}>Most Recent Month vs. Avg. of Prior Months</th>
            <th style={{ border: "1px solid #ccc", padding: "4px 8px" }}>Margin Risk Assessment</th>
            <th style={{ border: "1px solid #ccc", padding: "4px 8px" }}>Expense Growth Alignment</th>
            <th style={{ border: "1px solid #ccc", padding: "4px 8px" }}>Suggested Action</th>
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
            const sorted = [...data].sort((a, b) => {
              const aIdx = actionOrder.indexOf(a["Efficiency Alert"] || "")
              const bIdx = actionOrder.indexOf(b["Efficiency Alert"] || "")
              return (aIdx === -1 ? 999 : aIdx) - (bIdx === -1 ? 999 : bIdx)
            });
            return sorted.map((row, i) => (
              <tr key={i}>
                <td style={{ border: "1px solid #ccc", padding: "4px 8px" }}>{row["Category"]}</td>
                <td style={{ border: "1px solid #ccc", padding: "4px 8px" }}>{row["May"]}</td>
                <td style={{ border: "1px solid #ccc", padding: "4px 8px" }}>{row["Anchor vs Prior Avg ($)"]}</td>
                <td style={{ border: "1px solid #ccc", padding: "4px 8px" }}>{row["Margin Risk Assessment"]}</td>
                <td style={{ border: "1px solid #ccc", padding: "4px 8px" }}>{row["Expense Growth Alignment"]}</td>
                <td style={{ border: "1px solid #ccc", padding: "4px 8px" }}>{row["Efficiency Alert"]}</td>
              </tr>
            ))
          })()}
        </tbody>
      </table>
    </div>
  )
}

export default App
