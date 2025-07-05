# Stock Opname DSS (Decision Support System)

A React-based Decision Support System for stock opname (inventory management) with variance analysis capabilities.

## Features

- **Stock Opname Analysis**: Calculate and analyze inventory discrepancies
- **Variance Calculation**: Automatic variance analysis with severity classification
- **Interactive Dashboard**: Visual representation of inventory data
- **Excel Integration**: Import/export functionality for Excel files
- **Real-time Charts**: Dynamic visualization using Recharts library

## Technologies Used

- **Frontend**: React 19 + TypeScript
- **Build Tool**: Vite
- **Styling**: Tailwind CSS v4
- **Charts**: Recharts
- **Icons**: Lucide React
- **Excel Processing**: XLSX library

## Getting Started

### Prerequisites

- Node.js (version 16 or higher)
- npm or yarn

### Installation

1. Clone the repository:
```bash
git clone https://github.com/Sasamantiri13/StockOpname.git
cd StockOpname
```

2. Install dependencies:
```bash
npm install
```

3. Start the development server:
```bash
npm run dev
```

4. Open your browser and navigate to `http://localhost:5173`

## Available Scripts

- `npm run dev` - Start development server
- `npm run build` - Build for production
- `npm run preview` - Preview production build

## Project Structure

```
src/
├── main.tsx              # Application entry point
├── App.tsx               # Main App component
├── stock-opname-dss.tsx  # Main DSS component
├── stock-opname-simple.tsx # Simplified version
├── TestComponent.tsx     # Test component
├── style.css            # Global styles
└── vite-env.d.ts        # Vite type definitions
```

## Features Overview

### Variance Analysis
- Calculates variance between system and physical counts
- Severity classification (Low, Medium, High, Critical)
- Automatic variance percentage calculation

### Visual Dashboard
- Real-time charts showing inventory status
- Variance distribution visualization
- Summary statistics and KPIs

### Excel Integration
- Import stock data from Excel files
- Export analysis results
- Template generation for data entry

## Contributing

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add some amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

## License

This project is licensed under the MIT License.

## Author

- **Sasamantiri13** - [GitHub](https://github.com/Sasamantiri13)

## Acknowledgments

- Built with React and modern web technologies
- Inspired by inventory management best practices
- Uses open-source libraries for enhanced functionality
