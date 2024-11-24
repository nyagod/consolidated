import React from 'react';
import ProgressBar from './ProgressBar';  // Import ProgressBar component
import { Link } from 'react-router-dom';

function DyeingReports() {
    // Example data for total carpets
    const totalCarpets = {
        freshDyeing: 120,
        redyeing: 45,
        additionalDyeing: 30
    };

    return (
        <div className="dyeing-reports">
            <h2>Dyeing Reports</h2>

            {/* Add the ProgressBar component */}
            <ProgressBar />

            <div className="details-section">
                <h3>Carpet Dyeing Summary</h3>
                <ul className="carpet-summary-list">
                    <li>
                        <strong>Fresh Dyeing:</strong> {totalCarpets.freshDyeing} carpets
                    </li>
                    <li>
                        <strong>Redyeing:</strong> {totalCarpets.redyeing} carpets
                    </li>
                    <li>
                        <strong>Additional Dyeing:</strong> {totalCarpets.additionalDyeing} carpets
                    </li>
                </ul>

                <h3>Additional Details</h3>
                <p>Some details about the dyeing reports can go here.</p>

                {/* Example links for navigation */}
                <Link to="/more-details">View More Details</Link>
            </div>
        </div>
    );
}

export default DyeingReports;
