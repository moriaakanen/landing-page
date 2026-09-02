import { ChevronLeft, ChevronRight, MapPin } from 'lucide-react';
import { useState } from 'react';

export function AttendanceHistory() {
  const [selectedDate, setSelectedDate] = useState(23);

  const attendanceData = [
    {
      date: '23',
      day: 'Thu',
      checkIn: '07:58',
      checkOut: '-',
      totalHours: '-',
      location: 'Office, West Jakarta, Indonesia'
    },
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
      totalHours: '08:06',
      location: 'Office, West Jakarta, Indonesia'
    }
  ];

  const calendarDays = [
    { date: 29, month: 'prev', day: 'Sun' },
    { date: 30, month: 'prev', day: 'Mon' },
    { date: 31, month: 'prev', day: 'Tue' },
    { date: 1, month: 'current', day: 'Wed' },
    { date: 2, month: 'current', day: 'Thu' },
    { date: 3, month: 'current', day: 'Fri' },
    { date: 4, month: 'current', day: 'Sat' },
    { date: 5, month: 'current', day: 'Sun' },
    { date: 6, month: 'current', day: 'Mon' },
    { date: 7, month: 'current', day: 'Tue' },
    { date: 8, month: 'current', day: 'Wed' },
    { date: 9, month: 'current', day: 'Thu' },
    { date: 10, month: 'current', day: 'Fri' },
    { date: 11, month: 'current', day: 'Sat' },
    { date: 12, month: 'current', day: 'Sun' },
    { date: 13, month: 'current', day: 'Mon' },
    { date: 14, month: 'current', day: 'Tue' },
    { date: 15, month: 'current', day: 'Wed' },
    { date: 16, month: 'current', day: 'Thu' },
    { date: 17, month: 'current', day: 'Fri' },
    { date: 18, month: 'current', day: 'Sat' },
    { date: 19, month: 'current', day: 'Sun' },
    { date: 20, month: 'current', day: 'Mon' },
    { date: 21, month: 'current', day: 'Tue' },
    { date: 22, month: 'current', day: 'Wed' },
    { date: 23, month: 'current', day: 'Thu' },
    { date: 24, month: 'current', day: 'Fri' },
    { date: 25, month: 'current', day: 'Sat' },
    { date: 26, month: 'current', day: 'Sun' },
    { date: 27, month: 'current', day: 'Mon' },
    { date: 28, month: 'current', day: 'Tue' },
    { date: 29, month: 'current', day: 'Wed' },
    { date: 30, month: 'current', day: 'Thu' },
    { date: 1, month: 'next', day: 'Fri' },
    { date: 2, month: 'next', day: 'Sat' },
  ];

  return (
    <div className="p-6 pb-8">
      {/* Header */}
      <div className="flex items-center justify-between mb-6">
        <button className="p-2 hover:bg-gray-100 rounded-lg">
          <ChevronLeft className="w-5 h-5 text-gray-600" />
        </button>
        <h1 className="text-gray-900 text-lg">Attendance History</h1>
        <div className="w-9"></div>
      </div>

      {/* Month Selector */}
      <div className="flex items-center justify-center gap-4 mb-6">
        <button className="p-1 hover:bg-gray-100 rounded-full">
          <ChevronLeft className="w-5 h-5 text-[#5a7d6f]" />
        </button>
        <h2 className="text-gray-900 min-w-[150px] text-center">November 2023</h2>
        <button className="p-1 hover:bg-gray-100 rounded-full">
          <ChevronRight className="w-5 h-5 text-[#5a7d6f]" />
        </button>
      </div>

      {/* Calendar */}
      <div className="mb-6">
        <div className="grid grid-cols-7 gap-2 mb-3">
          {['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'].map((day) => (
            <div key={day} className="text-center text-xs text-gray-500 py-2">
              {day}
            </div>
          ))}
        </div>
        <div className="grid grid-cols-7 gap-2">
          {calendarDays.map((day, index) => (
            <button
              key={index}
              onClick={() => day.month === 'current' && setSelectedDate(day.date)}
              className={`
                aspect-square flex items-center justify-center rounded-lg text-sm
                ${day.month !== 'current' ? 'text-gray-300' : 'text-gray-900'}
                ${day.date === selectedDate && day.month === 'current' ? 'bg-[#5a7d6f] text-white' : 'hover:bg-gray-100'}
                ${[13, 14, 15, 16, 20, 21, 22].includes(day.date) && day.month === 'current' ? 'relative after:absolute after:bottom-1 after:w-1 after:h-1 after:bg-red-500 after:rounded-full' : ''}
              `}
            >
              {day.date}
            </button>
          ))}
        </div>
      </div>

      {/* Your Attendance */}
      <h3 className="text-gray-900 mb-4">Your Attendance</h3>

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
    </div>
  );
}
