import React, { useState } from 'react';
import { AreaChart, Area, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer, Legend } from 'recharts';

function DyeingMasters() {
    // Example data for the dyeing masters and their efficiency percentages
    const data = [
        { name: 'Master 1', efficiency: 20, efficiency2: 30, efficiency3: 30 },
        { name: 'Master 2', efficiency: 90, efficiency2: 92, efficiency3: 85 },
        { name: 'Master 3', efficiency: 180, efficiency2: 188, efficiency3: 278 },
        { name: 'Master 4', efficiency: 195, efficiency2: 297, efficiency3: 392 },
        { name: 'Master 5', efficiency: 288, efficiency2: 390, efficiency3: 489 },
    ];

    const [activeIndex, setActiveIndex] = useState(null);

    const handleAreaEnter = (index) => {
        setActiveIndex(index);
    };

    const handleAreaLeave = () => {
        setActiveIndex(null);
    };

    return (
        <div className="dyeing-masters">
            <h2>Dyeing Masters Efficiency Comparison</h2>
            <ResponsiveContainer width="100%" height={300}>
                <AreaChart data={data}>
                    <CartesianGrid strokeDasharray="3 3" />
                    <XAxis dataKey="name" />
                    <YAxis />
                    <Tooltip />
                    <Legend />
                    {/* Area for Master 1 */}
                    <Area
                        type="monotone"
                        dataKey="efficiency"
                        stroke="#536def"
                        fill="#536def"
                        fillOpacity={0.4}
                        onMouseEnter={handleAreaEnter}
                        onMouseLeave={handleAreaLeave}
                        strokeWidth={2}
                        name="Master 1"
                    />
                    {/* Area for Master 2 */}
                    <Area
                        type="monotone"
                        dataKey="efficiency2"
                        stroke="#ff8c00"
                        fill="#ff8c00"
                        fillOpacity={0.4}
                        onMouseEnter={handleAreaEnter}
                        onMouseLeave={handleAreaLeave}
                        strokeWidth={2}
                        name="Master 2"
                    />
                    {/* Area for Master 3 */}
                    <Area
                        type="monotone"
                        dataKey="efficiency3"
                        stroke="#4caf50"
                        fill="#4caf50"
                        fillOpacity={0.4}
                        onMouseEnter={handleAreaEnter}
                        onMouseLeave={handleAreaLeave}
                        strokeWidth={2}
                        name="Master 3"
                    />
                </AreaChart>
            </ResponsiveContainer>
        </div>
    );
}

export default DyeingMasters;
