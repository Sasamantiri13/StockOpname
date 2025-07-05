import React from 'react';

const TestComponent: React.FC = () => {
  return (
    <div className="min-h-screen bg-gray-50 p-4">
      <div className="max-w-screen-xl mx-auto">
        <h1 className="text-3xl font-bold text-gray-900 mb-4">
          Test Component - App is Working!
        </h1>
        <p className="text-lg text-gray-600">
          If you can see this, the basic React app is working correctly.
        </p>
      </div>
    </div>
  );
};

export default TestComponent;

