import { ChevronLeft, Calendar, Upload, X, FileText } from 'lucide-react';
import { useState } from 'react';

export function AbsenceRequest() {
  const [selectedType, setSelectedType] = useState('sick');
  const [startDate, setStartDate] = useState('');
  const [endDate, setEndDate] = useState('');
  const [reason, setReason] = useState('');
  const [attachment, setAttachment] = useState<string | null>(null);
  const [activeTab, setActiveTab] = useState('request');

  const absenceHistory = [
    {
      id: 1,
      type: 'Sick Leave',
      date: '15-17 Nov 2023',
      duration: '3 Days',
      status: 'Approved',
      reason: 'Flu and fever',
      submittedDate: '14 Nov 2023'
    },
    {
      id: 2,
      type: 'Personal Leave',
      date: '08 Nov 2023',
      duration: '1 Day',
      status: 'Approved',
      reason: 'Family matters',
      submittedDate: '06 Nov 2023'
    },
    {
      id: 3,
      type: 'Sick Leave',
      date: '25 Oct 2023',
      duration: '2 Days',
      status: 'Approved',
      reason: 'Medical check-up',
      submittedDate: '24 Oct 2023'
    },
    {
      id: 4,
      type: 'Emergency Leave',
      date: '10 Oct 2023',
      duration: '1 Day',
      status: 'Pending',
      reason: 'Emergency situation',
      submittedDate: '10 Oct 2023'
    }
  ];

  const absenceTypes = [
    { id: 'sick', label: 'Sick Leave', icon: '🤒', color: 'bg-red-50 text-red-700 border-red-200' },
    { id: 'personal', label: 'Personal Leave', icon: '👤', color: 'bg-blue-50 text-blue-700 border-blue-200' },
    { id: 'emergency', label: 'Emergency', icon: '🚨', color: 'bg-orange-50 text-orange-700 border-orange-200' },
    { id: 'other', label: 'Other', icon: '📋', color: 'bg-gray-50 text-gray-700 border-gray-200' }
  ];

  const handleFileUpload = () => {
    // Simulate file upload
    setAttachment('medical_certificate.pdf');
  };

  const handleSubmit = () => {
    // Handle form submission
    alert('Absence request submitted successfully!');
  };

  return (
    <div className="p-6 pb-8">
      {/* Header */}
      <div className="flex items-center justify-between mb-6">
        <button className="p-2 hover:bg-gray-100 rounded-lg">
          <ChevronLeft className="w-5 h-5 text-gray-600" />
        </button>
        <h1 className="text-gray-900 text-lg">Absence Request</h1>
        <div className="w-9"></div>
      </div>

      {/* Tabs */}
      <div className="flex gap-4 mb-6 border-b border-gray-200">
        <button
          onClick={() => setActiveTab('request')}
          className={`pb-3 border-b-2 transition-colors ${
            activeTab === 'request'
              ? 'border-[#5a7d6f] text-[#5a7d6f]'
              : 'border-transparent text-gray-400'
          }`}
        >
          New Request
        </button>
        <button
          onClick={() => setActiveTab('history')}
          className={`pb-3 border-b-2 transition-colors ${
            activeTab === 'history'
              ? 'border-[#5a7d6f] text-[#5a7d6f]'
              : 'border-transparent text-gray-400'
          }`}
        >
          History
        </button>
      </div>

      {activeTab === 'request' ? (
        <div>
          {/* Employee Info Card */}
          <div className="bg-[#5a7d6f] rounded-2xl p-5 mb-6 text-white">
            <div className="grid grid-cols-2 gap-4">
              <div>
                <p className="text-white/70 text-xs mb-1">Employee Name</p>
                <p className="text-sm">Akhmad Maariz</p>
              </div>
              <div>
                <p className="text-white/70 text-xs mb-1">Employee ID</p>
                <p className="text-sm">2023988231</p>
              </div>
              <div>
                <p className="text-white/70 text-xs mb-1">Position</p>
                <p className="text-sm">UI/UX Designer</p>
              </div>
              <div>
                <p className="text-white/70 text-xs mb-1">Department</p>
                <p className="text-sm">Design Team</p>
              </div>
            </div>
          </div>

          {/* Absence Type Selection */}
          <div className="mb-6">
            <p className="text-gray-900 mb-3">Type of Absence</p>
            <div className="grid grid-cols-2 gap-3">
              {absenceTypes.map((type) => (
                <button
                  key={type.id}
                  onClick={() => setSelectedType(type.id)}
                  className={`p-4 rounded-2xl border-2 transition-all ${
                    selectedType === type.id
                      ? 'border-[#5a7d6f] bg-[#5a7d6f]/5 shadow-sm'
                      : 'border-gray-200 bg-white'
                  }`}
                >
                  <div className="text-2xl mb-2">{type.icon}</div>
                  <p className={`text-sm ${selectedType === type.id ? 'text-[#5a7d6f]' : 'text-gray-700'}`}>
                    {type.label}
                  </p>
                </button>
              ))}
            </div>
          </div>

          {/* Date Range */}
          <div className="mb-6">
            <p className="text-gray-900 mb-3">Duration</p>
            <div className="grid grid-cols-2 gap-3">
              <div>
                <label className="text-xs text-gray-500 mb-2 block">Start Date</label>
                <div className="relative">
                  <Calendar className="absolute left-3 top-3 w-4 h-4 text-gray-400" />
                  <input
                    type="date"
                    value={startDate}
                    onChange={(e) => setStartDate(e.target.value)}
                    className="w-full pl-10 pr-3 py-3 border border-gray-200 rounded-xl bg-white text-gray-900"
                  />
                </div>
              </div>
              <div>
                <label className="text-xs text-gray-500 mb-2 block">End Date</label>
                <div className="relative">
                  <Calendar className="absolute left-3 top-3 w-4 h-4 text-gray-400" />
                  <input
                    type="date"
                    value={endDate}
                    onChange={(e) => setEndDate(e.target.value)}
                    className="w-full pl-10 pr-3 py-3 border border-gray-200 rounded-xl bg-white text-gray-900"
                  />
                </div>
              </div>
            </div>
            {startDate && endDate && (
              <div className="mt-3 bg-[#5a7d6f]/10 rounded-xl px-4 py-2">
                <p className="text-[#5a7d6f] text-sm">
                  Total Duration: {Math.ceil((new Date(endDate).getTime() - new Date(startDate).getTime()) / (1000 * 60 * 60 * 24)) + 1} Day(s)
                </p>
              </div>
            )}
          </div>

          {/* Reason */}
          <div className="mb-6">
            <p className="text-gray-900 mb-3">Reason for Absence</p>
            <textarea
              value={reason}
              onChange={(e) => setReason(e.target.value)}
              placeholder="Please describe the reason for your absence..."
              rows={4}
              className="w-full px-4 py-3 border border-gray-200 rounded-2xl bg-white text-gray-900 placeholder:text-gray-400 resize-none"
            />
          </div>

          {/* Attachment */}
          <div className="mb-6">
            <p className="text-gray-900 mb-3">Attachment (Optional)</p>
            {attachment ? (
              <div className="bg-white border border-gray-200 rounded-2xl p-4 flex items-center justify-between">
                <div className="flex items-center gap-3">
                  <div className="w-10 h-10 bg-[#5a7d6f]/10 rounded-lg flex items-center justify-center">
                    <FileText className="w-5 h-5 text-[#5a7d6f]" />
                  </div>
                  <div>
                    <p className="text-gray-900 text-sm">{attachment}</p>
                    <p className="text-gray-400 text-xs">Medical Certificate</p>
                  </div>
                </div>
                <button
                  onClick={() => setAttachment(null)}
                  className="p-2 hover:bg-gray-100 rounded-lg"
                >
                  <X className="w-4 h-4 text-gray-400" />
                </button>
              </div>
            ) : (
              <button
                onClick={handleFileUpload}
                className="w-full border-2 border-dashed border-gray-300 rounded-2xl p-6 hover:border-[#5a7d6f] hover:bg-[#5a7d6f]/5 transition-colors"
              >
                <Upload className="w-8 h-8 text-gray-400 mx-auto mb-2" />
                <p className="text-gray-600 text-sm mb-1">Upload Document</p>
                <p className="text-gray-400 text-xs">
                  Medical certificate, doctor's note, etc.
                </p>
              </button>
            )}
          </div>

          {/* Submit Button */}
          <button
            onClick={handleSubmit}
            className="w-full py-4 bg-[#5a7d6f] text-white rounded-full hover:bg-[#4a6d5f] transition-colors"
          >
            Submit Request
          </button>
        </div>
      ) : (
        /* History Tab */
        <div>
          {/* Summary Cards */}
          <div className="grid grid-cols-3 gap-3 mb-6">
            <div className="bg-white border border-gray-200 rounded-2xl p-4 text-center">
              <p className="text-gray-500 text-xs mb-1">Total</p>
              <p className="text-gray-900 text-2xl">7</p>
            </div>
            <div className="bg-green-50 border border-green-200 rounded-2xl p-4 text-center">
              <p className="text-green-700 text-xs mb-1">Approved</p>
              <p className="text-green-700 text-2xl">6</p>
            </div>
            <div className="bg-orange-50 border border-orange-200 rounded-2xl p-4 text-center">
              <p className="text-orange-700 text-xs mb-1">Pending</p>
              <p className="text-orange-700 text-2xl">1</p>
            </div>
          </div>

          {/* Filter */}
          <div className="flex gap-2 mb-6 overflow-x-auto">
            <button className="px-4 py-2 bg-[#5a7d6f] text-white rounded-full text-sm whitespace-nowrap">
              All
            </button>
            <button className="px-4 py-2 bg-white border border-gray-200 text-gray-900 rounded-full text-sm whitespace-nowrap">
              Sick Leave
            </button>
            <button className="px-4 py-2 bg-white border border-gray-200 text-gray-900 rounded-full text-sm whitespace-nowrap">
              Personal
            </button>
            <button className="px-4 py-2 bg-white border border-gray-200 text-gray-900 rounded-full text-sm whitespace-nowrap">
              Emergency
            </button>
          </div>

          {/* History List */}
          <div className="space-y-3">
            {absenceHistory.map((item) => (
              <div key={item.id} className="bg-white border border-gray-200 rounded-2xl p-4">
                <div className="flex items-start justify-between mb-3">
                  <div className="flex-1">
                    <div className="flex items-center gap-2 mb-1">
                      <h3 className="text-gray-900">{item.type}</h3>
                      <span
                        className={`px-2 py-1 rounded-full text-xs ${
                          item.status === 'Approved'
                            ? 'bg-green-100 text-green-700'
                            : item.status === 'Pending'
                            ? 'bg-orange-100 text-orange-700'
                            : 'bg-red-100 text-red-700'
                        }`}
                      >
                        {item.status}
                      </span>
                    </div>
                    <p className="text-gray-500 text-sm">{item.date}</p>
                  </div>
                  <div className="text-right">
                    <p className="text-[#5a7d6f]">{item.duration}</p>
                  </div>
                </div>
                <div className="bg-gray-50 rounded-xl p-3">
                  <p className="text-xs text-gray-500 mb-1">Reason:</p>
                  <p className="text-gray-900 text-sm">{item.reason}</p>
                </div>
                <p className="text-xs text-gray-400 mt-2">
                  Submitted: {item.submittedDate}
                </p>
              </div>
            ))}
          </div>
        </div>
      )}
    </div>
  );
}
