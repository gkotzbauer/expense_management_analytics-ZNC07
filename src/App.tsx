import { useState, useEffect } from "react"
import "./App.css"
import * as XLSX from "xlsx"

function App() {
  const [data, setData] = useState<any[][]>([])
  const [error, setError] = useState<string | null>(null)

  useEffect(() => {
    const fetchExcel = async () => {
      try {
        console.log("Starting to fetch Excel data...")
        
        // Fetch from generic Excel file in public folder
        const response = await fetch("/expense-analysis.xlsx")
        console.log("Response status:", response.status, response.ok)
        
        if (!response.ok) throw new Error("Failed to fetch Excel file")
        
        console.log("Successfully fetched file, parsing...")
        const blob = await response.blob()
        const arrayBuffer = await blob.arrayBuffer()
        const workbook = XLSX.read(arrayBuffer, { type: "array" })
        const sheetName = workbook.SheetNames[0]
        const worksheet = workbook.Sheets[sheetName]
        const jsonData: any[][] = XLSX.utils.sheet_to_json(worksheet, { header: 1 })
        console.log("Parsed data rows:", jsonData.length)
        setData(jsonData)
      } catch (err: any) {
        console.error("Error in fetchExcel:", err)
        setError(err.message || "Error loading data")
      }
    }
    fetchExcel()
  }, [])

  return (
    <div>
      <h1>Midlothian Margin Performance Dashboard</h1>
      {error && <div style={{ color: "red" }}>{error}</div>}
      <table style={{ margin: "0 auto", borderCollapse: "collapse" }}>
        <tbody>
          {data.map((row, i) => (
            <tr key={i}>
              {row.map((cell, j) => (
                <td key={j} style={{ border: "1px solid #ccc", padding: "4px 8px" }}>{cell}</td>
              ))}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  )
}

export default App
