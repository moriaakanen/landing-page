import { ChevronLeft, TrendingUp, TrendingDown, Clock, CheckCircle } from 'lucide-react';

export function Analytics() {
  const monthlyData = [
    { month: 'Jan', present: 20, late: 2, absent: 1 },
    { month: 'Feb', present: 18, late: 3, absent: 2 },
    { month: 'Mar', present: 22, late: 1, absent: 0 },
    { month: 'Apr', present: 19, late: 2, absent: 2 },
    { month: 'May', present: 21, late: 1, absent: 1 },
    { month: 'Jun', present: 20, late: 2, absent: 1 },
  ];

  return (
    <div className="p-6 pb-8">
      {/* Header */}
      <div className="flex items-center justify-between mb-6">
        <button className="p-2 hover:bg-gray-100 rounded-lg">
          <ChevronLeft className="w-5 h-5 text-gray-600" />
        </button>
        <h1 className="text-gray-900 text-lg">Analytics</h1>
        <div className="w-9"></div>
      </div>

      {/* Period Selector */}
      <div className="flex gap-2 mb-6">
        <button className="px-4 py-2 bg-[#5a7d6f] text-white rounded-full text-sm">
          This Month
        </button>
        <button className="px-4 py-2 bg-white border border-gray-200 text-gray-900 rounded-full text-sm">
          Last Month
        </button>
        <button className="px-4 py-2 bg-white border border-gray-200 text-gray-900 rounded-full text-sm">
          This Year
        </button>
      </div>

      {/* Summary Cards */}
      <div className="grid grid-cols-2 gap-4 mb-6">
        <div className="bg-white border border-gray-200 rounded-2xl p-4">
          <div className="flex items-center justify-between mb-2">
            <span className="text-gray-600 text-sm">Total Days</span>
            <CheckCircle className="w-4 h-4 text-[#5a7d6f]" />
          </div>
          <p className="text-gray-900 text-2xl mb-1">23</p>
          <div className="flex items-center gap-1 text-green-600 text-xs">
            <TrendingUp className="w-3 h-3" />
            <span>+5% from last month</span>
          </div>
        </div>

        <div className="bg-white border border-gray-200 rounded-2xl p-4">
          <div className="flex items-center justify-between mb-2">
            <span className="text-gray-600 text-sm">Present</span>
            <CheckCircle className="w-4 h-4 text-[#5a7d6f]" />
          </div>
          <p className="text-gray-900 text-2xl mb-1">20</p>
          <div className="flex items-center gap-1 text-green-600 text-xs">
            <TrendingUp className="w-3 h-3" />
            <span>+3% from last month</span>
          </div>
        </div>

        <div className="bg-white border border-gray-200 rounded-2xl p-4">
          <div className="flex items-center justify-between mb-2">
            <span className="text-gray-600 text-sm">Late</span>
            <Clock className="w-4 h-4 text-orange-500" />
          </div>
          <p className="text-gray-900 text-2xl mb-1">2</p>
          <div className="flex items-center gap-1 text-red-600 text-xs">
            <TrendingDown className="w-3 h-3" />
            <span>-1% from last month</span>
          </div>
        </div>

        <div className="bg-white border border-gray-200 rounded-2xl p-4">
          <div className="flex items-center justify-between mb-2">
            <span className="text-gray-600 text-sm">Absent</span>
            <Clock className="w-4 h-4 text-red-500" />
          </div>
          <p className="text-gray-900 text-2xl mb-1">1</p>
          <div className="flex items-center gap-1 text-green-600 text-xs">
            <TrendingUp className="w-3 h-3" />
            <span>Same as last month</span>
          </div>
        </div>
      </div>

      {/* Work Hours */}
      <div className="bg-[#5a7d6f] rounded-2xl p-5 mb-6">
        <div className="flex items-center justify-between mb-4">
          <h3 className="text-white">Average Work Hours</h3>
          <span className="text-white/70 text-sm">Nov 2023</span>
        </div>
        <div className="text-white text-4xl mb-2">8.2</div>
        <p className="text-white/70 text-sm">hours per day</p>
      </div>

      {/* Monthly Chart */}
      <div className="mb-6">
        <h3 className="text-gray-900 mb-4">Monthly Overview</h3>
        <div className="bg-white border border-gray-200 rounded-2xl p-4">
          {/* Simple bar chart representation */}
          <div className="flex items-end justify-between gap-2 h-40 mb-4">
            {monthlyData.map((data, index) => {
              const maxValue = 23;
              const height = (data.present / maxValue) * 100;
              return (
                <div key={index} className="flex-1 flex flex-col items-center gap-2">
                  <div className="w-full flex flex-col gap-1">
                    <div
                      className="w-full bg-[#5a7d6f] rounded-t"
                      style={{ height: `${height}%` }}
                    />
                    <div
                      className="w-full bg-orange-400 rounded"
                      style={{ height: `${(data.late / maxValue) * 100}%` }}
                    />
                    <div
                      className="w-full bg-red-400 rounded"
                      style={{ height: `${(data.absent / maxValue) * 100}%` }}
                    />
                  </div>
                </div>
              );
            })}
          </div>
          <div className="flex justify-between text-xs text-gray-500">
            {monthlyData.map((data, index) => (
              <span key={index}>{data.month}</span>
            ))}
          </div>
          
          {/* Legend */}
          <div className="flex gap-4 mt-4 justify-center">
            <div className="flex items-center gap-2">
              <div className="w-3 h-3 bg-[#5a7d6f] rounded-full" />
              <span className="text-xs text-gray-600">Present</span>
            </div>
            <div className="flex items-center gap-2">
              <div className="w-3 h-3 bg-orange-400 rounded-full" />
              <span className="text-xs text-gray-600">Late</span>
            </div>
            <div className="flex items-center gap-2">
              <div className="w-3 h-3 bg-red-400 rounded-full" />
              <span className="text-xs text-gray-600">Absent</span>
            </div>
          </div>
        </div>
      </div>

      {/* Performance Score */}
      <div className="bg-white border border-gray-200 rounded-2xl p-5">
        <h3 className="text-gray-900 mb-4">Attendance Score</h3>
        <div className="flex items-center justify-center mb-4">
          <div className="relative w-32 h-32">
            <svg className="w-full h-full transform -rotate-90">
              <circle
                cx="64"
                cy="64"
                r="56"
                fill="none"
                stroke="#e5e7eb"
                strokeWidth="12"
              />
              <circle
                cx="64"
                cy="64"
                r="56"
                fill="none"
                stroke="#5a7d6f"
                strokeWidth="12"
                strokeDasharray={`${2 * Math.PI * 56}`}
                strokeDashoffset={`${2 * Math.PI * 56 * (1 - 0.87)}`}
                strokeLinecap="round"
              />
            </svg>
            <div className="absolute inset-0 flex items-center justify-center flex-col">
              <span className="text-3xl text-gray-900">87%</span>
              <span className="text-xs text-gray-500">Excellent</span>
            </div>
          </div>
        </div>
        <p className="text-center text-gray-600 text-sm">
          Your attendance is better than 78% of employees
        </p>
      </div>
    </div>
  );
}
