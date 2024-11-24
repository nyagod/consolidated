import React, { useEffect, useState } from 'react';
import { BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer, Legend } from 'recharts';
import './YarnOverview.scss';

function YarnOverview() {
    const [transformedData, setTransformedData] = useState([]);
    const [chartData, setChartData] = useState([]);
    const [loading, setLoading] = useState(true);

    const roundToTwo = (num) => Math.round((num + Number.EPSILON) * 100) / 100;

    useEffect(() => {
        fetch('http://localhost:3000/dyed-store') // API endpoint
            .then((response) => response.json())
            .then((data) => {
                const yarnTypes = [...new Set(data.map(item => item.Type))];
                const weeks = [1, 2, 3, 4];

                // Prepare table data
                const tableData = yarnTypes.map(type => {
                    const row = { type };
                    weeks.forEach(week => {
                        const weekData = data.find(item => item.Type === type && item.Week === week);
                        row[`week${week}_received`] = roundToTwo(weekData?.Received || 0);
                        row[`week${week}_issued`] = roundToTwo(weekData?.Issued || 0);
                        row[`week${week}_returned`] = roundToTwo(weekData?.Returned || 0);
                    });
                    return row;
                });

                // Prepare chart data
                const chartData = weeks.map(week => {
                    let weekData = { name: `Week ${week}` };
                    yarnTypes.forEach(type => {
                        const typeData = data.find(d => d.Type === type && d.Week === week) || {};
                        weekData[`${type}_received`] = roundToTwo(typeData.Received || 0);
                        weekData[`${type}_issued`] = roundToTwo(typeData.Issued || 0);
                        weekData[`${type}_returned`] = roundToTwo(typeData.Returned || 0);
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
        <div className="yarn-overview">
            <h2>Yarn Overview Weekly Data</h2>

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
                                dataKey={`${type}_received`}
                                name={`${type} - Received`}
                                fill={yarnColors[type]}
                                barSize={25}
                                stackId={`stack${index}`}
                            />
                            <Bar
                                dataKey={`${type}_issued`}
                                name={`${type} - Issued`}
                                fill={yarnColors[type]}
                                barSize={25}
                                stackId={`stack${index}`}
                            />
                            <Bar
                                dataKey={`${type}_returned`}
                                name={`${type} - Returned`}
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
                            <th key={week} colSpan={3}>Week {week}</th>
                        ))}
                    </tr>
                    <tr>
                        <th></th>
                        {[1, 2, 3, 4].map(week => (
                            <>
                                <th key={`week${week}_received`}>Received</th>
                                <th key={`week${week}_issued`}>Issued</th>
                                <th key={`week${week}_returned`}>Returned</th>
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
                                        <td>{row[`week${week}_received`].toFixed(2)}</td>
                                        <td>{row[`week${week}_issued`].toFixed(2)}</td>
                                        <td>{row[`week${week}_returned`].toFixed(2)}</td>
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

export default YarnOverview;
