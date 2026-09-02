import { ChevronLeft, ChevronRight, User, Bell, Lock, Globe, HelpCircle, LogOut, Moon } from 'lucide-react';
import { useState } from 'react';

export function Settings() {
  const [notifications, setNotifications] = useState(true);
  const [darkMode, setDarkMode] = useState(false);

  const settingsSections = [
    {
      title: 'Account',
      items: [
        { icon: User, label: 'Profile Settings', action: 'navigate' },
        { icon: Lock, label: 'Privacy & Security', action: 'navigate' },
        { icon: Bell, label: 'Notifications', action: 'toggle', state: notifications, setState: setNotifications },
      ]
    },
    {
      title: 'Preferences',
      items: [
        { icon: Globe, label: 'Language', subtitle: 'English', action: 'navigate' },
        { icon: Moon, label: 'Dark Mode', action: 'toggle', state: darkMode, setState: setDarkMode },
      ]
    },
    {
      title: 'Support',
      items: [
        { icon: HelpCircle, label: 'Help Center', action: 'navigate' },
        { icon: User, label: 'About App', subtitle: 'Version 1.0.0', action: 'navigate' },
      ]
    }
  ];

  return (
    <div className="p-6 pb-8">
      {/* Header */}
      <div className="flex items-center justify-between mb-6">
        <button className="p-2 hover:bg-gray-100 rounded-lg">
          <ChevronLeft className="w-5 h-5 text-gray-600" />
        </button>
        <h1 className="text-gray-900 text-lg">Settings</h1>
        <div className="w-9"></div>
      </div>

      {/* Profile Card */}
      <div className="bg-[#5a7d6f] rounded-2xl p-5 mb-6 text-white">
        <div className="flex items-center gap-4">
          <div className="w-16 h-16 bg-white/20 rounded-full flex items-center justify-center">
            <User className="w-8 h-8" />
          </div>
          <div className="flex-1">
            <h2 className="text-lg mb-1">Akhmad Maariz</h2>
            <p className="text-white/70 text-sm">UI/UX Designer</p>
            <p className="text-white/70 text-xs mt-1">ID: 2023988231</p>
          </div>
          <button className="p-2 hover:bg-white/10 rounded-lg">
            <ChevronRight className="w-5 h-5" />
          </button>
        </div>
      </div>

      {/* Settings Sections */}
      {settingsSections.map((section, sectionIndex) => (
        <div key={sectionIndex} className="mb-6">
          <h3 className="text-gray-500 text-xs uppercase tracking-wider mb-3 px-2">
            {section.title}
          </h3>
          <div className="bg-white border border-gray-200 rounded-2xl overflow-hidden">
            {section.items.map((item, itemIndex) => {
              const Icon = item.icon;
              return (
                <div
                  key={itemIndex}
                  className={`flex items-center gap-4 p-4 ${
                    itemIndex !== section.items.length - 1 ? 'border-b border-gray-100' : ''
                  }`}
                >
                  <div className="w-10 h-10 bg-gray-100 rounded-full flex items-center justify-center">
                    <Icon className="w-5 h-5 text-[#5a7d6f]" />
                  </div>
                  <div className="flex-1">
                    <p className="text-gray-900">{item.label}</p>
                    {item.subtitle && (
                      <p className="text-gray-500 text-sm">{item.subtitle}</p>
                    )}
                  </div>
                  {item.action === 'navigate' && (
                    <ChevronRight className="w-5 h-5 text-gray-400" />
                  )}
                  {item.action === 'toggle' && item.setState && (
                    <button
                      onClick={() => item.setState(!item.state)}
                      className={`relative w-12 h-6 rounded-full transition-colors ${
                        item.state ? 'bg-[#5a7d6f]' : 'bg-gray-300'
                      }`}
                    >
                      <div
                        className={`absolute top-0.5 left-0.5 w-5 h-5 bg-white rounded-full transition-transform ${
                          item.state ? 'translate-x-6' : 'translate-x-0'
                        }`}
                      />
                    </button>
                  )}
                </div>
              );
            })}
          </div>
        </div>
      ))}

      {/* Logout Button */}
      <button className="w-full flex items-center justify-center gap-3 py-4 bg-red-50 text-red-600 rounded-2xl hover:bg-red-100 transition-colors">
        <LogOut className="w-5 h-5" />
        <span>Logout</span>
      </button>
    </div>
  );
}
