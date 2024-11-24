import React from 'react';
import { Pie, PieChart, ResponsiveContainer, Tooltip, Cell, Legend } from 'recharts';
import './ProgressBar.scss'; // Import your existing CSS

function ProgressBar() {
    // Data for the Pie Chart representing carpet counts
    const carpetData = [
        { name: 'Fresh Dyeing', value: 120 },
        { name: 'Redyeing', value: 45 },
        { name: 'Additional Dyeing', value: 30 },
    ];

    // Custom colors for each segment
    const COLORS = ['#536def', '#ff8c00', '#4caf50'];

    return (
        <div className="progress_bar">
            <div className="top">
                <p>Carpet Dyeing Summary</p>
            </div>

            <div className="middle">
                <div className="progress">
                    <ResponsiveContainer width="100%" height={300}>
                        <PieChart>
                            <Pie
                                dataKey="value"
                                data={carpetData}
                                cx="50%"
                                cy="50%"
                                outerRadius={80}
                                fill="#8884d8"
                                label={({ name, value }) => `${name}: ${value}`}
                                innerRadius={60} // Creates a donut chart effect
                            >
                                {carpetData.map((entry, index) => (
                                    <Cell key={`cell-${index}`} fill={COLORS[index % COLORS.length]} />
                                ))}
                            </Pie>
                            <Tooltip />
                            <Legend verticalAlign="bottom" height={36} />
                        </PieChart>
                    </ResponsiveContainer>
                </div>
                <p className="summary-text">Breakdown of carpet dyeing types.</p>
            </div>
        </div>
    );
}

export default ProgressBar;
