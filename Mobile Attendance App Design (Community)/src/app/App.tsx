import { useState } from 'react';
import { Home } from './components/Home';
import { AttendanceHistory } from './components/AttendanceHistory';
import { LeaveRequest } from './components/LeaveRequest';
import { AbsenceRequest } from './components/AbsenceRequest';
import { Analytics } from './components/Analytics';
import { Settings } from './components/Settings';
import { Navigation } from './components/Navigation';

export default function App() {
  const [currentPage, setCurrentPage] = useState('home');

  const renderPage = () => {
    switch (currentPage) {
      case 'home':
        return <Home />;
      case 'analytics':
        return <Analytics />;
      case 'attendance':
        return <AttendanceHistory />;
      case 'absence':
        return <AbsenceRequest />;
      case 'leave':
        return <LeaveRequest />;
      case 'settings':
        return <Settings />;
      default:
        return <Home />;
    }
  };

  return (
    <div className="min-h-screen bg-gray-50 pb-20">
      <div className="max-w-md mx-auto bg-white min-h-screen shadow-xl">
        {renderPage()}
        <Navigation currentPage={currentPage} onPageChange={setCurrentPage} />
      </div>
    </div>
  );
}