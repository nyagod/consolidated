import React from 'react';
import DyeingReports from './DyeingReports';
import DyeingMasters from './DyeingMasters';
import YarnIssued from './YarnIssued';
import TotalSqmtOfCarpets from './TotalSqmtOfCarpets';
import Available from './Available';
import Totalcolors from './totalcolors';
import { Link } from 'react-router-dom';
import './Dashboard.scss'; // Import your custom CSS for dashboard

function Dashboard() {
    return (
        <div className="dashboard">
            <h1>Dashboard</h1>

            {/* KPI Cards or Summary Section */}
            <div className="kpi-summary">
                <div className="kpi-item">
                    <h3>Dyeing Reports</h3>
                    <DyeingReports />
                </div>
                <div className="kpi-item">
                    <h3>Dyeing Masters</h3>
                    <DyeingMasters />
                </div>
                <div className="kpi-item">
                    <h3>Yarn Issued</h3>
                    <YarnIssued />
                </div>
                <div className="kpi-item">
                    <h3>Total Sqmt of Carpets</h3>
                    <TotalSqmtOfCarpets />
                </div>
                <div className="kpi-item">
                    <h3>Available</h3>
                    <Available />
                </div>
                <div className="kpi-item">
                    <h3>Total Colors</h3>
                    <Totalcolors />
                </div>
            </div>

            {/* Navigation Links to Each Section (Optional) */}
            <div className="dashboard-links">
                <Link to="/dyeing-reports">View Dyeing Reports</Link>
                <Link to="/dyeing-masters">View Dyeing Masters</Link>
                <Link to="/yarn-issued">View Yarn Issued</Link>
                <Link to="/total-sqmt-of-carpets">View Total Sqmt of Carpets</Link>
                <Link to="/available">View Available</Link>
            </div>
        </div>
    );
}

export default Dashboard;
