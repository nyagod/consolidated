import React, { useEffect, useState } from 'react';
import { BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer, Legend } from 'recharts';
import './Available.scss';

function Available() {
    const [transformedData, setTransformedData] = useState([]);
    const [chartData, setChartData] = useState([]);
    const [loading, setLoading] = useState(true);

    useEffect(() => {
        fetch('http://localhost:3000/available-section') // Updated API endpoint for available yarn
            .then((response) => response.json())
            .then((data) => {
                const yarnTypes = [...new Set(data.map(item => item.Type))];
                const weeks = [1, 2, 3, 4]; // Define the 4 weeks

                // Prepare table data (Ensure null values are replaced with 0)
                const tableData = yarnTypes.map(type => {
                    const row = { type };
                    weeks.forEach(week => {
                        const weekData = data.find(item => item.Type === type && item.Week === week);
                        row[`week${week}_in`] = weekData?.In || 0; // Replace null with 0
                        row[`week${week}_out`] = weekData?.Out || 0; // Replace null with 0
                    });
                    return row;
                });

                // Prepare chart data: One entry per week, each yarn type gets its own "in" and "out"
                const chartData = weeks.map(week => {
                    let weekData = { name: `Week ${week}` };
                    yarnTypes.forEach(type => {
                        const typeData = data.find(d => d.Type === type && d.Week === week) || {};
                        weekData[`${type}_in`] = typeData.In || 0;  // Replace null with 0
                        weekData[`${type}_out`] = typeData.Out || 0; // Replace null with 0
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
        <div className="available">
            <h2>Available Yarn Weekly Data</h2>

            {/* Bar Chart */}
            <ResponsiveContainer width="100%" height={400}>
                <BarChart
                    data={chartData}
                    margin={{ top: 20, right: 30, left: 0, bottom: 20 }}
                    barCategoryGap="10%"
                >
                    <CartesianGrid strokeDasharray="3 3" />
                    <XAxis dataKey="name" />
                    <YAxis />
                    <Tooltip />
                    <Legend />
                    {Object.keys(yarnColors).map((type, index) => (
                        <React.Fragment key={type}>
                            <Bar
                                dataKey={`${type}_in`}
                                name={`${type} - In`}
                                fill={yarnColors[type]}
                                barSize={25}
                                stackId={`stack${index}`}
                            />
                            <Bar
                                dataKey={`${type}_out`}
                                name={`${type} - Out`}
                                fill={yarnColors[type]}
                                barSize={25}
                                stackId={`stack${index}`}
                            />
                        </React.Fragment>
                    ))}
                </BarChart>
            </ResponsiveContainer>

            {/* Data Table */}
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
                                        <td>{row[`week${week}_in`].toFixed(2)}</td>
                                        <td>{row[`week${week}_out`].toFixed(2)}</td>
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

export default Available;
