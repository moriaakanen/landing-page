import { ChevronLeft, ChevronDown, Calendar } from 'lucide-react';
import { useState } from 'react';

export function LeaveRequest() {
  const [leaveType, setLeaveType] = useState('12 Day');
  const [selectedLeaveCategory, setSelectedLeaveCategory] = useState('Annual Leave');

  return (
    <div className="p-6 pb-8">
      {/* Header */}
      <div className="flex items-center justify-between mb-6">
        <button className="p-2 hover:bg-gray-100 rounded-lg">
          <ChevronLeft className="w-5 h-5 text-gray-600" />
        </button>
        <h1 className="text-gray-900 text-lg">Leave</h1>
        <div className="w-9"></div>
      </div>

      {/* Employee Info */}
      <div className="grid grid-cols-2 gap-4 mb-6">
        <div>
          <p className="text-xs text-gray-500 mb-1">Employee Name</p>
          <p className="text-gray-900">Akhmad Maariz</p>
        </div>
        <div>
          <p className="text-xs text-gray-500 mb-1">Employee ID</p>
          <p className="text-gray-900">2023988231</p>
        </div>
        <div>
          <p className="text-xs text-gray-500 mb-1">Job Position</p>
          <p className="text-gray-900">UI/UX Designer</p>
        </div>
        <div>
          <p className="text-xs text-gray-500 mb-1">Type</p>
          <p className="text-gray-900">Full Time</p>
        </div>
      </div>

      {/* Leave Type Selector */}
      <div className="grid grid-cols-2 gap-3 mb-6">
        <button
          onClick={() => setLeaveType('12 Day')}
          className={`py-3 rounded-xl border ${
            leaveType === '12 Day'
              ? 'bg-[#5a7d6f] text-white border-[#5a7d6f]'
              : 'bg-white text-gray-900 border-gray-200'
          }`}
        >
          <p className="text-sm">12 Day</p>
          <p className="text-xs opacity-70">Available Leave</p>
        </button>
        <button
          onClick={() => setLeaveType('8 Day')}
          className={`py-3 rounded-xl border ${
            leaveType === '8 Day'
              ? 'bg-[#5a7d6f] text-white border-[#5a7d6f]'
              : 'bg-white text-gray-900 border-gray-200'
          }`}
        >
          <p className="text-sm">8 Day</p>
          <p className="text-xs opacity-70">Used Leave</p>
        </button>
      </div>

      {/* Tabs */}
      <div className="flex gap-4 mb-6 border-b border-gray-200">
        <button className="pb-3 border-b-2 border-[#5a7d6f] text-[#5a7d6f]">
          Leave Request
        </button>
        <button className="pb-3 text-gray-400">
          Leave History
        </button>
      </div>

      {/* Duration */}
      <div className="mb-6">
        <p className="text-gray-900 mb-3">Duration</p>
        <div className="grid grid-cols-2 gap-3">
          <div className="relative">
            <Calendar className="absolute left-3 top-3 w-4 h-4 text-gray-400" />
            <input
              type="text"
              value="27/11/2023"
              readOnly
              className="w-full pl-10 pr-3 py-3 border border-gray-200 rounded-xl bg-white text-gray-900"
            />
          </div>
          <div className="relative">
            <span className="absolute left-1/2 top-1/2 -translate-x-1/2 -translate-y-1/2 text-gray-400 text-sm">
              to
            </span>
          </div>
          <div className="relative col-start-1">
            <Calendar className="absolute left-3 top-3 w-4 h-4 text-gray-400" />
            <input
              type="text"
              value="30/11/2023"
              readOnly
              className="w-full pl-10 pr-3 py-3 border border-gray-200 rounded-xl bg-white text-gray-900"
            />
          </div>
        </div>
      </div>

      {/* Types of Leave */}
      <div className="mb-6">
        <p className="text-gray-900 mb-3">Types of leave</p>
        <div className="relative">
          <select
            value={selectedLeaveCategory}
            onChange={(e) => setSelectedLeaveCategory(e.target.value)}
            className="w-full px-4 py-3 border border-gray-200 rounded-xl bg-white text-gray-900 appearance-none cursor-pointer"
          >
            <option>Annual Leave</option>
            <option>Sick Leave</option>
            <option>Personal Leave</option>
            <option>Emergency Leave</option>
          </select>
          <ChevronDown className="absolute right-4 top-1/2 -translate-y-1/2 w-4 h-4 text-gray-400 pointer-events-none" />
        </div>
      </div>

      {/* Reason */}
      <div className="mb-6">
        <p className="text-gray-900 mb-3">Reason (Optional)</p>
        <textarea
          placeholder="Write description here..."
          rows={4}
          className="w-full px-4 py-3 border border-gray-200 rounded-xl bg-gray-50 text-gray-900 placeholder:text-gray-400 resize-none"
        />
      </div>

      {/* Submit Button */}
      <button className="w-full py-4 bg-[#5a7d6f] text-white rounded-full hover:bg-[#4a6d5f] transition-colors">
        Submit Request
      </button>
    </div>
  );
}
