import React, { useEffect, useState } from 'react';
import { BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer, Legend } from 'recharts';
import './YarnIssued.scss';

function YarnIssued() {
    const [transformedData, setTransformedData] = useState([]); // Transformed data for the table
    const [chartData, setChartData] = useState([]); // Data for the bar chart
    const [loading, setLoading] = useState(true); // State for loading status

    useEffect(() => {
        fetch('http://localhost:3000/undyed-yarn')  // API endpoint
            .then((response) => response.json())
            .then((data) => {
                const yarnTypes = [...new Set(data.map(item => item.Type))];
                const weeks = [1, 2, 3, 4];  // Defining the 4 weeks

                // Prepare table data (Ensure null values are replaced with 0 and round to 2 decimal places)
                const tableData = yarnTypes.map(type => {
                    const row = { type };
                    weeks.forEach(week => {
                        const weekData = data.find(item => item.Type === type && item.Week === week);
                        row[`week${week}_in`] = weekData?.In ? weekData.In.toFixed(2) : '0.00'; // Replace null with 0 and round to 2 decimals
                        row[`week${week}_out`] = weekData?.Out ? weekData.Out.toFixed(2) : '0.00'; // Replace null with 0 and round to 2 decimals
                    });
                    return row;
                });

                // Prepare chart data: One entry per week, each yarn type gets its own "in" and "out"
                const chartData = weeks.map(week => {
                    let weekData = { name: `Week ${week}` };
                    yarnTypes.forEach(type => {
                        const typeData = data.find(d => d.Type === type && d.Week === week) || {};
                        weekData[`${type}_in`] = typeData.In ? typeData.In.toFixed(2) : '0.00';  // Replace null with 0 and round to 2 decimals
                        weekData[`${type}_out`] = typeData.Out ? typeData.Out.toFixed(2) : '0.00'; // Replace null with 0 and round to 2 decimals
                    });
                    return weekData;
                });

                setTransformedData(tableData);
                setChartData(chartData);
                setLoading(false);
            })
            .catch((error) => {
                console.error('Error fetching data: ', error);
                setLoading(false);
            });
    }, []);

    if (loading) {
        return <div>Loading...</div>;
    }

    const yarnColors = {
        Yarn: '#536def',
        Silk: '#f57c00',
        Cotton: '#4caf50',
    };

    return (
        <div className="yarn-issued">
            <h2>Yarn Issued Weekly Data</h2>

            <ResponsiveContainer width="100%" height={400}>
                <BarChart
                    data={chartData}
                    margin={{ top: 20, right: 30, left: 0, bottom: 20 }}
                    barCategoryGap="10%" // Reduce spacing between groups for better alignment with thicker bars
                >
                    <CartesianGrid strokeDasharray="3 3" />
                    <XAxis dataKey="name" />
                    <YAxis />
                    <Tooltip />
                    <Legend />
                    {/* Dynamically map each yarn type */}
                    {Object.keys(yarnColors).map((type, index) => (
                        <React.Fragment key={type}>
                            {/* "Issued" bars for the type */}
                            <Bar
                                dataKey={`${type}_in`}
                                name={`${type} - Issued`}
                                fill={yarnColors[type]}
                                barSize={25} // Thicker bars
                                stackId={`stack${index}`} // Group bars by yarn type but separate in/out
                            />
                            {/* "Received" bars for the type */}
                            <Bar
                                dataKey={`${type}_out`}
                                name={`${type} - Received`}
                                fill={yarnColors[type]}
                                barSize={25} // Thicker bars
                                stackId={`stack${index}`} // Group bars by yarn type but separate in/out
                            />
                        </React.Fragment>
                    ))}
                </BarChart>
            </ResponsiveContainer>

            {/* Table */}
            <table>
                <thead>
                    <tr>
                        <th>Yarn Type</th>
                        {[1, 2, 3, 4].map(week => (
                            <th key={week} colSpan={2}>Week {week}</th>
                        ))}
                    </tr>
                    <tr>
                        <th></th>
                        {[1, 2, 3, 4].map(week => (
                            <>
                                <th key={`week${week}_in`}>In</th>
                                <th key={`week${week}_out`}>Out</th>
                            </>
                        ))}
                    </tr>
                </thead>
                <tbody>
                    {transformedData.map((row, index) => (
                        <tr key={index}>
                            <td>{row.type}</td>
                            {Array.from({ length: 4 }, (_, weekIndex) => {
                                const week = weekIndex + 1;
                                return (
                                    <React.Fragment key={week}>
                                        <td>{parseFloat(row[`week${week}_in`]).toFixed(2)}</td>
                                        <td>{parseFloat(row[`week${week}_out`]).toFixed(2)}</td>
                                    </React.Fragment>
                                );
                            })}
                        </tr>
                    ))}
                </tbody>
            </table>
        </div>
    );
}

export default YarnIssued;
