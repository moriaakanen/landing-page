import { Home, BarChart3, Calendar, LogOut, Settings, FileX } from 'lucide-react';

interface NavigationProps {
  currentPage: string;
  onPageChange: (page: string) => void;
}

export function Navigation({ currentPage, onPageChange }: NavigationProps) {
  const navItems = [
    { id: 'home', label: 'Home', icon: Home },
    { id: 'analytics', label: 'Analytics', icon: BarChart3 },
    { id: 'attendance', label: 'Attendance', icon: Calendar },
    { id: 'absence', label: 'Absence', icon: FileX },
    { id: 'leave', label: 'Leave', icon: LogOut },
    { id: 'settings', label: 'Settings', icon: Settings },
  ];

  return (
    <nav className="fixed bottom-0 left-0 right-0 bg-white border-t border-gray-200 max-w-md mx-auto shadow-lg">
      <div className="flex justify-around items-center h-16">
        {navItems.map((item) => {
          const Icon = item.icon;
          const isActive = currentPage === item.id;
          
          return (
            <button
              key={item.id}
              onClick={() => onPageChange(item.id)}
              className="flex flex-col items-center justify-center flex-1 h-full relative"
            >
              {isActive && item.id === 'attendance' && (
                <div className="absolute -top-2 w-14 h-14 bg-[#5a7d6f] rounded-full flex items-center justify-center shadow-lg">
                  <Icon className="w-6 h-6 text-white" />
                </div>
              )}
              {(!isActive || item.id !== 'attendance') && (
                <>
                  <Icon className={`w-5 h-5 ${isActive ? 'text-[#5a7d6f]' : 'text-gray-400'}`} />
                  <span className={`text-xs mt-1 ${isActive ? 'text-[#5a7d6f]' : 'text-gray-400'}`}>
                    {item.label}
                  </span>
                </>
              )}
            </button>
          );
        })}
      </div>
    </nav>
  );
}