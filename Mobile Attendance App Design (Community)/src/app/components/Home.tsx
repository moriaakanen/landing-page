import { Bell, MapPin, CheckCircle, Clock, Camera, MapPinned } from 'lucide-react';
import { useState } from 'react';

export function Home() {
  const [isCheckedIn, setIsCheckedIn] = useState(true);
  const [checkInTime, setCheckInTime] = useState('07:58');
  const [checkOutTime, setCheckOutTime] = useState('');
  const [showCheckInModal, setShowCheckInModal] = useState(false);
  const [showCheckOutModal, setShowCheckOutModal] = useState(false);

  const handleCheckIn = () => {
    const now = new Date();
    const time = `${String(now.getHours()).padStart(2, '0')}:${String(now.getMinutes()).padStart(2, '0')}`;
    setCheckInTime(time);
    setIsCheckedIn(true);
    setShowCheckInModal(false);
  };

  const handleCheckOut = () => {
    const now = new Date();
    const time = `${String(now.getHours()).padStart(2, '0')}:${String(now.getMinutes()).padStart(2, '0')}`;
    setCheckOutTime(time);
    setShowCheckOutModal(false);
  };

  const attendanceData = [
    {
      date: '22',
      day: 'Wed',
      checkIn: '07:57',
      checkOut: '17:00',
      totalHours: '08:03',
      location: 'Office, West Jakarta, Indonesia'
    },
    {
      date: '21',
      day: 'Tue',
      checkIn: '08:03',
      checkOut: '17:08',
      totalHours: '08:05',
      location: 'Office, West Jakarta, Indonesia'
    },
    {
      date: '20',
      day: 'Mon',
      checkIn: '07:59',
      checkOut: '17:00',
      totalHours: '08:01',
      location: 'Office, West Jakarta, Indonesia'
    }
  ];

  return (
    <div className="p-6 pb-8">
      {/* Header */}
      <div className="flex justify-between items-start mb-6">
        <div>
          <p className="text-gray-500 text-sm">Good Morning,</p>
          <h1 className="text-gray-900 text-xl mt-1">Akhmad Maariz</h1>
        </div>
        <button className="p-2 hover:bg-gray-100 rounded-lg">
          <Bell className="w-5 h-5 text-gray-600" />
        </button>
      </div>

      {/* Location Badge */}
      <div className="bg-[#5a7d6f] text-white px-4 py-2 rounded-full inline-flex items-center gap-2 mb-2">
        <MapPin className="w-4 h-4" />
        <span className="text-sm">West Jakarta, Indonesia</span>
      </div>
      <p className="text-xs text-gray-500 mb-6">Thursday, 23 November 2023</p>

      {/* Check In/Out Buttons - New Prominent Section */}
      <div className="bg-gradient-to-br from-[#5a7d6f] to-[#4a6d5f] rounded-3xl p-6 mb-6 shadow-lg">
        <div className="text-center mb-6">
          <p className="text-white/80 text-sm mb-1">Current Time</p>
          <p className="text-white text-3xl">{new Date().toLocaleTimeString('en-US', { hour: '2-digit', minute: '2-digit' })}</p>
        </div>
        
        <div className="grid grid-cols-2 gap-4">
          <button
            onClick={() => setShowCheckInModal(true)}
            disabled={isCheckedIn && checkInTime}
            className="bg-white text-[#5a7d6f] py-4 rounded-2xl flex flex-col items-center gap-2 hover:bg-gray-50 transition-all disabled:opacity-50 disabled:cursor-not-allowed"
          >
            <CheckCircle className="w-6 h-6" />
            <span className="font-medium">Check In</span>
            {checkInTime && <span className="text-xs">{checkInTime}</span>}
          </button>
          
          <button
            onClick={() => setShowCheckOutModal(true)}
            disabled={!isCheckedIn || checkOutTime}
            className="bg-white/20 text-white py-4 rounded-2xl flex flex-col items-center gap-2 hover:bg-white/30 transition-all disabled:opacity-50 disabled:cursor-not-allowed backdrop-blur-sm"
          >
            <Clock className="w-6 h-6" />
            <span className="font-medium">Check Out</span>
            {checkOutTime && <span className="text-xs">{checkOutTime}</span>}
          </button>
        </div>
      </div>

      {/* Check In/Out Status Cards */}
      <div className="grid grid-cols-2 gap-4 mb-6">
        <div className="bg-white border border-gray-200 rounded-2xl p-4">
          <div className="flex items-center gap-2 mb-2">
            <CheckCircle className="w-4 h-4 text-[#5a7d6f]" />
            <span className="text-gray-600 text-sm">Check In</span>
          </div>
          <p className="text-gray-900 text-2xl">{checkInTime || '--:--'}</p>
          <p className="text-gray-400 text-xs mt-1">{checkInTime ? 'On Time' : 'Not checked in'}</p>
        </div>

        <div className="bg-white border border-gray-200 rounded-2xl p-4">
          <div className="flex items-center gap-2 mb-2">
            <Clock className={`w-4 h-4 ${checkOutTime ? 'text-[#5a7d6f]' : 'text-gray-400'}`} />
            <span className="text-gray-600 text-sm">Check Out</span>
          </div>
          <p className={`text-2xl ${checkOutTime ? 'text-gray-900' : 'text-gray-400'}`}>{checkOutTime || '--:--'}</p>
          <p className="text-gray-400 text-xs mt-1">{checkOutTime ? 'Completed' : 'Not checked out'}</p>
        </div>
      </div>

      {/* Stats */}
      <div className="grid grid-cols-2 gap-4 mb-6">
        <div className="bg-white border border-gray-200 rounded-2xl p-4">
          <div className="flex items-center gap-2 mb-2">
            <CheckCircle className="w-4 h-4 text-[#5a7d6f]" />
            <span className="text-gray-600 text-sm">Absence</span>
          </div>
          <p className="text-gray-900 text-2xl">3</p>
          <p className="text-gray-400 text-xs mt-1">Day</p>
        </div>

        <div className="bg-white border border-gray-200 rounded-2xl p-4">
          <div className="flex items-center gap-2 mb-2">
            <CheckCircle className="w-4 h-4 text-[#5a7d6f]" />
            <span className="text-gray-600 text-sm">Total Attended</span>
          </div>
          <p className="text-gray-900 text-2xl">15</p>
          <p className="text-gray-400 text-xs mt-1">Day</p>
        </div>
      </div>

      {/* Attendance History */}
      <div className="flex justify-between items-center mb-4">
        <h2 className="text-gray-900">Attendance History</h2>
        <button className="text-[#5a7d6f] text-sm">See More</button>
      </div>

      <div className="space-y-3">
        {attendanceData.map((item, index) => (
          <div key={index} className="bg-[#5a7d6f] rounded-2xl p-4 text-white">
            <div className="flex gap-4">
              <div className="bg-white/20 rounded-xl px-3 py-2 text-center min-w-[60px]">
                <p className="text-2xl">{item.date}</p>
                <p className="text-sm mt-1">{item.day}</p>
              </div>
              <div className="flex-1">
                <div className="grid grid-cols-3 gap-2 mb-2">
                  <div>
                    <p className="text-xs text-white/70">Check In</p>
                    <p className="text-sm">{item.checkIn}</p>
                  </div>
                  <div>
                    <p className="text-xs text-white/70">Check out</p>
                    <p className="text-sm">{item.checkOut}</p>
                  </div>
                  <div>
                    <p className="text-xs text-white/70">Total Hours</p>
                    <p className="text-sm">{item.totalHours}</p>
                  </div>
                </div>
                <div className="flex items-center gap-1 text-xs">
                  <MapPin className="w-3 h-3" />
                  <span>{item.location}</span>
                </div>
              </div>
            </div>
          </div>
        ))}
      </div>

      {/* Check In Modal */}
      {showCheckInModal && (
        <div className="fixed inset-0 bg-black/50 flex items-end justify-center z-50 max-w-md mx-auto">
          <div className="bg-white rounded-t-3xl w-full p-6 animate-slide-up">
            <div className="w-12 h-1 bg-gray-300 rounded-full mx-auto mb-6"></div>
            <h2 className="text-xl text-gray-900 mb-6 text-center">Check In</h2>
            
            <div className="space-y-4 mb-6">
              <div className="bg-gray-50 rounded-2xl p-4">
                <div className="flex items-center gap-3 mb-3">
                  <div className="w-10 h-10 bg-[#5a7d6f]/10 rounded-full flex items-center justify-center">
                    <Clock className="w-5 h-5 text-[#5a7d6f]" />
                  </div>
                  <div>
                    <p className="text-xs text-gray-500">Time</p>
                    <p className="text-gray-900">{new Date().toLocaleTimeString()}</p>
                  </div>
                </div>
                
                <div className="flex items-center gap-3">
                  <div className="w-10 h-10 bg-[#5a7d6f]/10 rounded-full flex items-center justify-center">
                    <MapPinned className="w-5 h-5 text-[#5a7d6f]" />
                  </div>
                  <div>
                    <p className="text-xs text-gray-500">Location</p>
                    <p className="text-gray-900 text-sm">West Jakarta, Indonesia</p>
                  </div>
                </div>
              </div>

              <div className="bg-gray-50 rounded-2xl p-4 border-2 border-dashed border-gray-300">
                <div className="flex flex-col items-center justify-center py-6">
                  <Camera className="w-12 h-12 text-gray-400 mb-2" />
                  <p className="text-gray-600 text-sm mb-1">Take a selfie</p>
                  <p className="text-gray-400 text-xs">Optional verification</p>
                </div>
              </div>

              <textarea
                placeholder="Add note (optional)..."
                rows={3}
                className="w-full px-4 py-3 border border-gray-200 rounded-2xl bg-white text-gray-900 placeholder:text-gray-400 resize-none"
              />
            </div>

            <div className="grid grid-cols-2 gap-3">
              <button
                onClick={() => setShowCheckInModal(false)}
                className="py-3 border border-gray-200 text-gray-600 rounded-full hover:bg-gray-50"
              >
                Cancel
              </button>
              <button
                onClick={handleCheckIn}
                className="py-3 bg-[#5a7d6f] text-white rounded-full hover:bg-[#4a6d5f]"
              >
                Confirm Check In
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Check Out Modal */}
      {showCheckOutModal && (
        <div className="fixed inset-0 bg-black/50 flex items-end justify-center z-50 max-w-md mx-auto">
          <div className="bg-white rounded-t-3xl w-full p-6 animate-slide-up">
            <div className="w-12 h-1 bg-gray-300 rounded-full mx-auto mb-6"></div>
            <h2 className="text-xl text-gray-900 mb-6 text-center">Check Out</h2>
            
            <div className="space-y-4 mb-6">
              <div className="bg-gray-50 rounded-2xl p-4">
                <div className="flex items-center gap-3 mb-3">
                  <div className="w-10 h-10 bg-[#5a7d6f]/10 rounded-full flex items-center justify-center">
                    <Clock className="w-5 h-5 text-[#5a7d6f]" />
                  </div>
                  <div>
                    <p className="text-xs text-gray-500">Time</p>
                    <p className="text-gray-900">{new Date().toLocaleTimeString()}</p>
                  </div>
                </div>
                
                <div className="flex items-center gap-3 mb-3">
                  <div className="w-10 h-10 bg-[#5a7d6f]/10 rounded-full flex items-center justify-center">
                    <CheckCircle className="w-5 h-5 text-[#5a7d6f]" />
                  </div>
                  <div>
                    <p className="text-xs text-gray-500">Checked In At</p>
                    <p className="text-gray-900">{checkInTime}</p>
                  </div>
                </div>

                <div className="flex items-center gap-3">
                  <div className="w-10 h-10 bg-[#5a7d6f]/10 rounded-full flex items-center justify-center">
                    <MapPinned className="w-5 h-5 text-[#5a7d6f]" />
                  </div>
                  <div>
                    <p className="text-xs text-gray-500">Location</p>
                    <p className="text-gray-900 text-sm">West Jakarta, Indonesia</p>
                  </div>
                </div>
              </div>

              <textarea
                placeholder="Add note about your work today (optional)..."
                rows={3}
                className="w-full px-4 py-3 border border-gray-200 rounded-2xl bg-white text-gray-900 placeholder:text-gray-400 resize-none"
              />
            </div>

            <div className="grid grid-cols-2 gap-3">
              <button
                onClick={() => setShowCheckOutModal(false)}
                className="py-3 border border-gray-200 text-gray-600 rounded-full hover:bg-gray-50"
              >
                Cancel
              </button>
              <button
                onClick={handleCheckOut}
                className="py-3 bg-[#5a7d6f] text-white rounded-full hover:bg-[#4a6d5f]"
              >
                Confirm Check Out
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
