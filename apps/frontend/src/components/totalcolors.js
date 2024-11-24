import React, { useEffect, useState } from 'react';
import { BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer } from 'recharts';
import './TotalColors.scss'; // Import the SCSS file

const TotalColors = () => {
  // State to hold the fetched color data
  const [colorData, setColorData] = useState([]);
  const [totalColorData, setTotalColorData] = useState(null); // State for the total color data

  // Fetch the data from the API on component mount
  useEffect(() => {
    fetch('http://localhost:3000/total-color-used') // Assuming the API is running locally on this endpoint
      .then(response => response.json())
      .then(data => {
        // Filter out the total summary data and store separately
        const totalData = data.find(item => item["COLORS"] === "TOTAL COLOR USED");
        const validData = data.filter(item => item["COLORS"] !== "TOTAL COLOR USED");
        
        // Process the valid data into a format suitable for the chart and table
        const processedData = validData.map(item => ({
          color: item["COLORS"], // Use 'COLORS' as the color label
          total: parseFloat(item["Total COST "] || 0).toFixed(2), // Round 'Total COST ' to 2 decimals
          costPerKg: parseFloat(item[" Cost per kg"] || 0).toFixed(2), // Add 'Cost per kg' value
          inGram: parseFloat(item["in GRAM"] || 0).toFixed(2), // Add 'in GRAM' value
        }));

        // Set valid data for the chart
        setColorData(processedData);

        // Set the total color data separately for display
        if (totalData) {
          setTotalColorData({
            totalColorUsed: parseFloat(totalData["Total COST "] || 0).toFixed(2),
            totalCost: parseFloat(totalData["Total COST "] || 0).toFixed(2),
          });
        }
      })
      .catch(error => {
        console.error('Error fetching the color data:', error);
      });
  }, []); // Empty dependency array ensures the effect runs only once when the component mounts

  // Calculate the maximum value for Y-axis domain
  const maxTotal = Math.max(...colorData.map(item => parseFloat(item.total)));

  return (
    <div className="total-colors-container">
      <h2>Total Colors Used</h2>
      
      <ResponsiveContainer width="100%" height={400}>
        <BarChart data={colorData} margin={{ top: 20, right: 30, left: 10, bottom: 80 }}>
          <CartesianGrid strokeDasharray="3 3" />
          <XAxis
            dataKey="color"
            tick={{ angle: -45, textAnchor: 'end', fontSize: 16 }}
            height={80}  // Increased space for better label display // Show all ticks
          />
          <YAxis domain={[0, maxTotal * 1.2]} />  {/* Dynamically set max value to 20% above the highest value */}
          <Tooltip />
          {/* <Legend /> */}
          <Bar dataKey="total" fill="#8884d8" barSize={50} />  {/* Increased bar size */}
        </BarChart>
      </ResponsiveContainer>

      {/* Table for valid data */}
      <table className="color-table">
        <thead>
          <tr>
            <th>Color</th>
            <th>Total Used</th>
            <th>Cost per kg</th>
            <th>In GRAM</th>
          </tr>
        </thead>
        <tbody>
          {colorData.map((color, index) => (
            <tr key={index}>
              <td>{color.color}</td>
              <td>{color.total}</td>
              <td>{color.costPerKg}</td>
              <td>{color.inGram}</td>
            </tr>
          ))}
        </tbody>
      </table>

      {/* Display Total Color Used as a separate summary section */}
      {totalColorData && (
        <div className="total-color-summary">
          <h3>Total Color Used</h3>
          <p><strong>Total Cost of All Colors: </strong>{totalColorData.totalColorUsed}</p>
        </div>
      )}
    </div>
  );
};

export default TotalColors;
