import React, { useState, useEffect } from 'react';
import { Area, AreaChart, CartesianGrid, Tooltip, XAxis, YAxis } from 'recharts';
import './chart.scss';

function Chart({ height, title }) {
    const [data, setData] = useState([]);

    useEffect(() => {
        fetch('http://localhost:3000/total-order-received') // Replace with your endpoint
            .then((response) => response.json())
            .then((rawData) => {
                // Filter and transform the data
                const cleanedData = rawData
                    .filter((row) => row["Party Name"] && row["Party Name"] !== "TOTAL ") // Ignore after TOTAL
                    .map((row) => ({
                        name: row["Party Name"],
                        FS: row["F Total Sq.mt"] || 0,
                        CS: row["C Total Sq. mt."] || 0,
                    }));
                setData(cleanedData);
            })
            .catch((error) => console.error('Error fetching chart data:', error));
    }, []);

    // Calculate total FS and CS, and total combined sum
    const totalFS = data.reduce((sum, row) => sum + row.FS, 0).toFixed(2);
    const totalCS = data.reduce((sum, row) => sum + row.CS, 0).toFixed(2);
    const totalCombined = (parseFloat(totalFS) + parseFloat(totalCS)).toFixed(2);

    return (
        <div className="chart_sec">
            {/* Chart Section */}
            <div className="title">
                <p>{title} (Last Month)</p>
            </div>
            <div style={{ width: '100%', height: 300 }}>
                <AreaChart
                    width={850}
                    height={height}
                    data={data}
                    margin={{ top: 10, right: 30, left: 0, bottom: 0 }}
                >
                    <defs>
                        <linearGradient id="FS" x1="0" y1="0" x2="0" y2="1">
                            <stop offset="5%" stopColor="#536def" stopOpacity={0.9} />
                            <stop offset="95%" stopColor="#536def" stopOpacity={0} />
                        </linearGradient>
                        <linearGradient id="CS" x1="0" y1="0" x2="0" y2="1">
                            <stop offset="5%" stopColor="#f57c00" stopOpacity={0.9} />
                            <stop offset="95%" stopColor="#f57c00" stopOpacity={0} />
                        </linearGradient>
                    </defs>
                    <XAxis dataKey="name" stroke="gray" />
                    <YAxis />
                    <CartesianGrid strokeDasharray="3 3" className="strokee" />
                    <Tooltip />
                    {/* FS Area */}
                    <Area
                        type="monotone"
                        dataKey="FS"
                        stroke="#536def"
                        fillOpacity={1}
                        fill="url(#FS)"
                        name="F Total Sq.mt"
                    />
                    {/* CS Area */}
                    <Area
                        type="monotone"
                        dataKey="CS"
                        stroke="#f57c00"
                        fillOpacity={1}
                        fill="url(#CS)"
                        name="C Total Sq.mt"
                    />
                </AreaChart>
            </div>

            {/* Table Section */}
            <div className="table_sec" style={{ marginTop: '20px' }}>
                <table>
                    <thead>
                        <tr>
                            <th>Party Name</th>
                            <th>F Total Sq.mt</th>
                            <th>C Total Sq.mt</th>
                        </tr>
                    </thead>
                    <tbody>
                        {data.map((row, index) => (
                            <tr key={index}>
                                <td>{row.name}</td>
                                <td>{row.FS}</td>
                                <td>{row.CS}</td>
                            </tr>
                        ))}
                    </tbody>
                </table>
            </div>

            {/* Total Sum Section */}
            <div className="total_sum" style={{ marginTop: '20px' }}>
                <p><strong>Total F Total Sq.mt: </strong>{totalFS}</p>
                <p><strong>Total C Total Sq.mt: </strong>{totalCS}</p>
                <p><strong>Total Combined Sq.mt: </strong>{totalCombined}</p> {/* Combined Total */}
            </div>
        </div>
    );
}

export default Chart;
