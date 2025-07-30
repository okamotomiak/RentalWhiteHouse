// FinancialManager.gs - Complete Financial Management System
// Handles all financial operations, reporting, and analysis

const FinancialManager = {
  
  /**
   * Column indexes for Budget sheet (1-based)
   */
  COL: {
    DATE: 1,
    TYPE: 2,
    DESCRIPTION: 3,
    AMOUNT: 4,
    CATEGORY: 5,
    PAYMENT_METHOD: 6,
    REFERENCE_NUMBER: 7,
    TENANT_GUEST: 8,
    RECEIPT: 9
  },
  
  /**
   * Log a payment or expense transaction
   */
  logPayment: function(transactionData) {
    try {
      const data = [
        transactionData.date || new Date(),
        transactionData.type || 'Other',
        transactionData.description || '',
        transactionData.amount || 0,
        transactionData.category || 'Other',
        transactionData.paymentMethod || '',
        transactionData.reference || '',
        transactionData.tenant || '',
        transactionData.receipt || ''
      ];
      
      SheetManager.addRow(CONFIG.SHEETS.BUDGET, data);
      Logger.log(`Financial transaction logged: ${transactionData.type} - ${Utils.formatCurrency(transactionData.amount)}`);
      
    } catch (error) {
      Logger.log(`Error logging payment: ${error.toString()}`);
      throw error;
    }
  },
  
  /**
   * Generate monthly financial report
   */
  generateMonthlyFinancialReport: function() {
    try {
      const now = new Date();
      const monthStart = new Date(now.getFullYear(), now.getMonth(), 1);
      const monthEnd = new Date(now.getFullYear(), now.getMonth() + 1, 0);
      const monthYear = Utils.formatDate(now, 'MMMM yyyy');
      
      const data = SheetManager.getAllData(CONFIG.SHEETS.BUDGET);
      const report = this.analyzeFinancialData(data, monthStart, monthEnd);
      
      const html = HtmlService.createHtmlOutput(`
        <div style="font-family: Arial, sans-serif; padding: 20px; line-height: 1.6;">
          <h2>💰 Monthly Financial Report - ${monthYear}</h2>
          
          <div style="display: grid; grid-template-columns: repeat(3, 1fr); gap: 20px; margin-bottom: 30px;">
            <div style="background: #e8f5e8; padding: 15px; border-radius: 8px; text-align: center;">
              <h3 style="margin: 0; color: #2e7d32;">Total Income</h3>
              <p style="font-size: 24px; margin: 5px 0; font-weight: bold;">${Utils.formatCurrency(report.totalIncome)}</p>
              <small>${report.incomeTransactions} transactions</small>
            </div>
            <div style="background: #ffebee; padding: 15px; border-radius: 8px; text-align: center;">
              <h3 style="margin: 0; color: #c62828;">Total Expenses</h3>
              <p style="font-size: 24px; margin: 5px 0; font-weight: bold;">${Utils.formatCurrency(report.totalExpenses)}</p>
              <small>${report.expenseTransactions} transactions</small>
            </div>
            <div style="background: #e3f2fd; padding: 15px; border-radius: 8px; text-align: center;">
              <h3 style="margin: 0; color: #1565c0;">Net Profit</h3>
              <p style="font-size: 24px; margin: 5px 0; font-weight: bold; color: ${report.netProfit >= 0 ? '#2e7d32' : '#c62828'};">${Utils.formatCurrency(report.netProfit)}</p>
              <small>Profit Margin: ${report.profitMargin}%</small>
            </div>
          </div>
          
          <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 30px;">
            <div>
              <h3>📈 Income Breakdown</h3>
              <div style="background: #f5f5f5; padding: 15px; border-radius: 8px;">
                ${Object.entries(report.incomeByCategory).map(([category, amount]) => `
                  <div style="display: flex; justify-content: space-between; margin-bottom: 8px;">
                    <span><strong>${category}:</strong></span>
                    <span>${Utils.formatCurrency(amount)}</span>
                  </div>
                `).join('')}
              </div>
            </div>
            
            <div>
              <h3>📉 Expense Breakdown</h3>
              <div style="background: #f5f5f5; padding: 15px; border-radius: 8px;">
                ${Object.entries(report.expensesByCategory).map(([category, amount]) => `
                  <div style="display: flex; justify-content: space-between; margin-bottom: 8px;">
                    <span><strong>${category}:</strong></span>
                    <span>${Utils.formatCurrency(amount)}</span>
                  </div>
                `).join('')}
              </div>
            </div>
          </div>
          
          <h3>💳 Payment Methods</h3>
          <div style="background: #f5f5f5; padding: 15px; border-radius: 8px; margin-bottom: 20px;">
            ${Object.entries(report.paymentMethods).map(([method, amount]) => `
              <div style="display: flex; justify-content: space-between; margin-bottom: 8px;">
                <span>${method || 'Not Specified'}:</span>
                <span>${Utils.formatCurrency(amount)}</span>
              </div>
            `).join('')}
          </div>
          
          <h3>📊 Key Metrics</h3>
          <div style="background: #f5f5f5; padding: 15px; border-radius: 8px;">
            <p><strong>Average Transaction Size:</strong> ${Utils.formatCurrency(report.avgTransactionSize)}</p>
            <p><strong>Largest Income:</strong> ${Utils.formatCurrency(report.largestIncome)} (${report.largestIncomeSource})</p>
            <p><strong>Largest Expense:</strong> ${Utils.formatCurrency(report.largestExpense)} (${report.largestExpenseSource})</p>
            <p><strong>Cash Flow Trend:</strong> ${report.cashFlowTrend}</p>
          </div>
          
          <div style="margin-top: 30px; text-align: center;">
            <button onclick="google.script.run.generateTaxReport()" style="margin: 5px; padding: 10px 20px;">Generate Tax Report</button>
            <button onclick="google.script.run.exportFinancialData()" style="margin: 5px; padding: 10px 20px;">Export Data</button>
          </div>
        </div>
      `)
        .setWidth(900)
        .setHeight(700);
      
      SpreadsheetApp.getUi().showModalDialog(html, 'Monthly Financial Report');
      
    } catch (error) {
      handleSystemError(error, 'generateMonthlyFinancialReport');
    }
  },
  
  /**
   * Analyze financial data for a given period
   */
  analyzeFinancialData: function(data, startDate, endDate) {
    const analysis = {
      totalIncome: 0,
      totalExpenses: 0,
      netProfit: 0,
      profitMargin: 0,
      incomeTransactions: 0,
      expenseTransactions: 0,
      avgTransactionSize: 0,
      incomeByCategory: {},
      expensesByCategory: {},
      paymentMethods: {},
      largestIncome: 0,
      largestIncomeSource: '',
      largestExpense: 0,
      largestExpenseSource: '',
      cashFlowTrend: 'Stable'
    };
    
    let totalTransactions = 0;
    let totalTransactionValue = 0;
    
    data.forEach(row => {
      const date = new Date(row[this.COL.DATE - 1]);
      const amount = row[this.COL.AMOUNT - 1] || 0;
      const category = row[this.COL.CATEGORY - 1] || 'Other';
      const description = row[this.COL.DESCRIPTION - 1] || '';
      const paymentMethod = row[this.COL.PAYMENT_METHOD - 1] || 'Not Specified';
      
      if (date >= startDate && date <= endDate) {
        totalTransactions++;
        totalTransactionValue += Math.abs(amount);
        
        // Track payment methods
        analysis.paymentMethods[paymentMethod] = (analysis.paymentMethods[paymentMethod] || 0) + Math.abs(amount);
        
        if (amount > 0) {
          // Income
          analysis.totalIncome += amount;
          analysis.incomeTransactions++;
          analysis.incomeByCategory[category] = (analysis.incomeByCategory[category] || 0) + amount;
          
          if (amount > analysis.largestIncome) {
            analysis.largestIncome = amount;
            analysis.largestIncomeSource = description;
          }
        } else if (amount < 0) {
          // Expense
          const expenseAmount = Math.abs(amount);
          analysis.totalExpenses += expenseAmount;
          analysis.expenseTransactions++;
          analysis.expensesByCategory[category] = (analysis.expensesByCategory[category] || 0) + expenseAmount;
          
          if (expenseAmount > analysis.largestExpense) {
            analysis.largestExpense = expenseAmount;
            analysis.largestExpenseSource = description;
          }
        }
      }
    });
    
    // Calculate derived metrics
    analysis.netProfit = analysis.totalIncome - analysis.totalExpenses;
    analysis.profitMargin = analysis.totalIncome > 0 ? 
      Math.round((analysis.netProfit / analysis.totalIncome) * 100) : 0;
    analysis.avgTransactionSize = totalTransactions > 0 ? 
      totalTransactionValue / totalTransactions : 0;
    
    // Determine cash flow trend (simplified)
    if (analysis.netProfit > analysis.totalIncome * 0.2) {
      analysis.cashFlowTrend = 'Strong Growth';
    } else if (analysis.netProfit > 0) {
      analysis.cashFlowTrend = 'Positive';
    } else if (analysis.netProfit > -analysis.totalIncome * 0.1) {
      analysis.cashFlowTrend = 'Break-even';
    } else {
      analysis.cashFlowTrend = 'Loss';
    }
    
    return analysis;
  },
  
  /**
   * Show revenue analysis dashboard
   */
  showRevenueAnalysis: function() {
    try {
      const data = SheetManager.getAllData(CONFIG.SHEETS.BUDGET);
      const now = new Date();
      
      // Current month
      const currentMonth = new Date(now.getFullYear(), now.getMonth(), 1);
      const currentAnalysis = this.analyzeFinancialData(data, currentMonth, now);
      
      // Previous month
      const previousMonth = new Date(now.getFullYear(), now.getMonth() - 1, 1);
      const previousMonthEnd = new Date(now.getFullYear(), now.getMonth(), 0);
      const previousAnalysis = this.analyzeFinancialData(data, previousMonth, previousMonthEnd);
      
      // Year to date
      const yearStart = new Date(now.getFullYear(), 0, 1);
      const ytdAnalysis = this.analyzeFinancialData(data, yearStart, now);
      
      // Calculate month-over-month changes
      const incomeChange = previousAnalysis.totalIncome > 0 ? 
        ((currentAnalysis.totalIncome - previousAnalysis.totalIncome) / previousAnalysis.totalIncome * 100) : 0;
      
      const html = HtmlService.createHtmlOutput(`
        <div style="font-family: Arial, sans-serif; padding: 20px; line-height: 1.6;">
          <h2>📈 Revenue Analysis Dashboard</h2>
          
          <div style="display: grid; grid-template-columns: repeat(3, 1fr); gap: 20px; margin-bottom: 30px;">
            <div style="background: #e8f5e8; padding: 15px; border-radius: 8px; text-align: center;">
              <h4 style="margin: 0; color: #2e7d32;">Current Month</h4>
              <p style="font-size: 20px; margin: 5px 0; font-weight: bold;">${Utils.formatCurrency(currentAnalysis.totalIncome)}</p>
              <small style="color: ${incomeChange >= 0 ? '#2e7d32' : '#c62828'};">
                ${incomeChange >= 0 ? '↗' : '↘'} ${Math.abs(incomeChange).toFixed(1)}% vs last month
              </small>
            </div>
            <div style="background: #fff3e0; padding: 15px; border-radius: 8px; text-align: center;">
              <h4 style="margin: 0; color: #ef6c00;">Previous Month</h4>
              <p style="font-size: 20px; margin: 5px 0; font-weight: bold;">${Utils.formatCurrency(previousAnalysis.totalIncome)}</p>
              <small>Net: ${Utils.formatCurrency(previousAnalysis.netProfit)}</small>
            </div>
            <div style="background: #e3f2fd; padding: 15px; border-radius: 8px; text-align: center;">
              <h4 style="margin: 0; color: #1565c0;">Year to Date</h4>
              <p style="font-size: 20px; margin: 5px 0; font-weight: bold;">${Utils.formatCurrency(ytdAnalysis.totalIncome)}</p>
              <small>Avg Monthly: ${Utils.formatCurrency(ytdAnalysis.totalIncome / (now.getMonth() + 1))}</small>
            </div>
          </div>
          
          <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 30px;">
            <div>
              <h3>📊 Revenue Streams (YTD)</h3>
              <div style="background: #f5f5f5; padding: 15px; border-radius: 8px;">
                ${Object.entries(ytdAnalysis.incomeByCategory)
                  .sort(([,a], [,b]) => b - a)
                  .map(([category, amount]) => {
                    const percentage = (amount / ytdAnalysis.totalIncome * 100).toFixed(1);
                    return `
                      <div style="margin-bottom: 10px;">
                        <div style="display: flex; justify-content: space-between; margin-bottom: 2px;">
                          <span><strong>${category}</strong></span>
                          <span>${Utils.formatCurrency(amount)} (${percentage}%)</span>
                        </div>
                        <div style="background: #ddd; height: 8px; border-radius: 4px;">
                          <div style="background: #4caf50; height: 100%; width: ${percentage}%; border-radius: 4px;"></div>
                        </div>
                      </div>
                    `;
                  }).join('')}
              </div>
            </div>
            
            <div>
              <h3>💸 Expense Categories (YTD)</h3>
              <div style="background: #f5f5f5; padding: 15px; border-radius: 8px;">
                ${Object.entries(ytdAnalysis.expensesByCategory)
                  .sort(([,a], [,b]) => b - a)
                  .map(([category, amount]) => {
                    const percentage = (amount / ytdAnalysis.totalExpenses * 100).toFixed(1);
                    return `
                      <div style="margin-bottom: 10px;">
                        <div style="display: flex; justify-content: space-between; margin-bottom: 2px;">
                          <span><strong>${category}</strong></span>
                          <span>${Utils.formatCurrency(amount)} (${percentage}%)</span>
                        </div>
                        <div style="background: #ddd; height: 8px; border-radius: 4px;">
                          <div style="background: #f44336; height: 100%; width: ${percentage}%; border-radius: 4px;"></div>
                        </div>
                      </div>
                    `;
                  }).join('')}
              </div>
            </div>
          </div>
          
          <h3>🎯 Performance Indicators</h3>
          <div style="display: grid; grid-template-columns: repeat(4, 1fr); gap: 15px; margin-bottom: 20px;">
            <div style="background: #f5f5f5; padding: 15px; border-radius: 8px; text-align: center;">
              <h4 style="margin: 0;">Profit Margin</h4>
              <p style="font-size: 18px; margin: 5px 0; font-weight: bold; color: ${ytdAnalysis.profitMargin >= 20 ? '#2e7d32' : ytdAnalysis.profitMargin >= 10 ? '#ef6c00' : '#c62828'};">
                ${ytdAnalysis.profitMargin}%
              </p>
            </div>
            <div style="background: #f5f5f5; padding: 15px; border-radius: 8px; text-align: center;">
              <h4 style="margin: 0;">Expense Ratio</h4>
              <p style="font-size: 18px; margin: 5px 0; font-weight: bold;">
                ${(ytdAnalysis.totalExpenses / ytdAnalysis.totalIncome * 100).toFixed(1)}%
              </p>
            </div>
            <div style="background: #f5f5f5; padding: 15px; border-radius: 8px; text-align: center;">
              <h4 style="margin: 0;">Monthly Growth</h4>
              <p style="font-size: 18px; margin: 5px 0; font-weight: bold; color: ${incomeChange >= 0 ? '#2e7d32' : '#c62828'};">
                ${incomeChange >= 0 ? '+' : ''}${incomeChange.toFixed(1)}%
              </p>
            </div>
            <div style="background: #f5f5f5; padding: 15px; border-radius: 8px; text-align: center;">
              <h4 style="margin: 0;">Cash Flow</h4>
              <p style="font-size: 18px; margin: 5px 0; font-weight: bold;">
                ${ytdAnalysis.cashFlowTrend}
              </p>
            </div>
          </div>
          
          <div style="background: #e3f2fd; padding: 15px; border-radius: 8px; margin-bottom: 20px;">
            <h4>💡 Insights & Recommendations</h4>
            ${this.generateFinancialInsights(currentAnalysis, previousAnalysis, ytdAnalysis)}
          </div>
          
          <div style="text-align: center;">
            <button onclick="google.script.run.showOccupancyAnalytics()" style="margin: 5px; padding: 10px 20px;">Occupancy Analytics</button>
            <button onclick="google.script.run.showProfitabilityDashboard()" style="margin: 5px; padding: 10px 20px;">Profitability Analysis</button>
          </div>
        </div>
      `)
        .setWidth(1000)
        .setHeight(800);
      
      SpreadsheetApp.getUi().showModalDialog(html, 'Revenue Analysis Dashboard');
      
    } catch (error) {
      handleSystemError(error, 'showRevenueAnalysis');
    }
  },
  
  /**
   * Generate financial insights and recommendations
   */
  generateFinancialInsights: function(current, previous, ytd) {
    const insights = [];
    
    // Revenue insights
    if (current.totalIncome > previous.totalIncome * 1.1) {
      insights.push('📈 <strong>Strong revenue growth</strong> this month. Consider investing in property improvements.');
    } else if (current.totalIncome < previous.totalIncome * 0.9) {
      insights.push('📉 <strong>Revenue declined</strong> from last month. Review pricing strategy and occupancy rates.');
    }
    
    // Profit margin insights
    if (ytd.profitMargin > 25) {
      insights.push('💰 <strong>Excellent profit margins.</strong> Consider expanding or improving amenities.');
    } else if (ytd.profitMargin < 10) {
      insights.push('⚠️ <strong>Low profit margins.</strong> Review expenses and consider rent increases.');
    }
    
    // Expense insights
    const expenseRatio = ytd.totalExpenses / ytd.totalIncome;
    if (expenseRatio > 0.7) {
      insights.push('💸 <strong>High expense ratio.</strong> Look for cost reduction opportunities.');
    }
    
    // Category-specific insights
    const maintenanceExpenses = ytd.expensesByCategory['Maintenance'] || 0;
    if (maintenanceExpenses > ytd.totalIncome * 0.15) {
      insights.push('🔧 <strong>High maintenance costs.</strong> Consider preventive maintenance programs.');
    }
    
    // Growth insights
    const monthlyAverage = ytd.totalIncome / (new Date().getMonth() + 1);
    if (current.totalIncome > monthlyAverage * 1.2) {
      insights.push('🚀 <strong>Above-average performance</strong> this month. Great work!');
    }
    
    if (insights.length === 0) {
      insights.push('📊 <strong>Stable performance.</strong> Continue monitoring key metrics for optimization opportunities.');
    }
    
    return '<ul><li>' + insights.join('</li><li>') + '</li></ul>';
  },
  
  /**
   * Generate tax report
   */
  generateTaxReport: function() {
    try {
      const now = new Date();
      const taxYear = now.getFullYear();
      const yearStart = new Date(taxYear, 0, 1);
      const yearEnd = new Date(taxYear, 11, 31);
      
      const data = SheetManager.getAllData(CONFIG.SHEETS.BUDGET);
      const taxData = this.prepareTaxData(data, yearStart, yearEnd);
      
      const html = HtmlService.createHtmlOutput(`
        <div style="font-family: Arial, sans-serif; padding: 20px; line-height: 1.6;">
          <h2>📋 Tax Report - ${taxYear}</h2>
          
          <div style="background: #fff3e0; padding: 15px; border-radius: 8px; margin-bottom: 20px;">
            <h4>⚠️ Important Tax Disclaimer</h4>
            <p>This report is for informational purposes only. Please consult with a qualified tax professional for official tax preparation and advice.</p>
          </div>
          
          <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 30px; margin-bottom: 30px;">
            <div>
              <h3 style="color: #2e7d32;">📈 Rental Income</h3>
              <div style="background: #e8f5e8; padding: 15px; border-radius: 8px;">
                <p><strong>Total Rental Income:</strong> ${Utils.formatCurrency(taxData.totalIncome)}</p>
                <hr style="margin: 10px 0;">
                ${Object.entries(taxData.incomeByCategory).map(([category, amount]) => `
                  <div style="display: flex; justify-content: space-between; margin-bottom: 5px;">
                    <span>${category}:</span>
                    <span>${Utils.formatCurrency(amount)}</span>
                  </div>
                `).join('')}
              </div>
            </div>
            
            <div>
              <h3 style="color: #c62828;">📉 Deductible Expenses</h3>
              <div style="background: #ffebee; padding: 15px; border-radius: 8px;">
                <p><strong>Total Deductions:</strong> ${Utils.formatCurrency(taxData.totalDeductions)}</p>
                <hr style="margin: 10px 0;">
                ${Object.entries(taxData.deductibleExpenses).map(([category, amount]) => `
                  <div style="display: flex; justify-content: space-between; margin-bottom: 5px;">
                    <span>${category}:</span>
                    <span>${Utils.formatCurrency(amount)}</span>
                  </div>
                `).join('')}
              </div>
            </div>
          </div>
          
          <div style="background: #e3f2fd; padding: 20px; border-radius: 8px; margin-bottom: 20px;">
            <h3>📊 Tax Summary</h3>
            <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 20px;">
              <div>
                <p><strong>Gross Rental Income:</strong> ${Utils.formatCurrency(taxData.totalIncome)}</p>
                <p><strong>Less: Total Expenses:</strong> ${Utils.formatCurrency(taxData.totalDeductions)}</p>
                <hr>
                <p style="font-size: 18px;"><strong>Net Rental Income:</strong> 
                  <span style="color: ${taxData.netIncome >= 0 ? '#2e7d32' : '#c62828'};">
                    ${Utils.formatCurrency(taxData.netIncome)}
                  </span>
                </p>
              </div>
              <div>
                <p><strong>Number of Properties:</strong> 1</p>
                <p><strong>Business Use %:</strong> 100%</p>
                <p><strong>Tax Year:</strong> ${taxYear}</p>
                <p><strong>Report Generated:</strong> ${Utils.formatDate(new Date())}</p>
              </div>
            </div>
          </div>
          
          <div style="background: #f5f5f5; padding: 15px; border-radius: 8px; margin-bottom: 20px;">
            <h4>📝 Common Tax Forms for Rental Property</h4>
            <ul>
              <li><strong>Schedule E (Form 1040):</strong> Supplemental Income and Loss</li>
              <li><strong>Form 4562:</strong> Depreciation and Amortization (if applicable)</li>
              <li><strong>Form 8825:</strong> Rental Real Estate Income (for partnerships)</li>
            </ul>
          </div>
          
          <div style="background: #f5f5f5; padding: 15px; border-radius: 8px;">
            <h4>💡 Tax Planning Tips</h4>
            <ul>
              <li>Keep detailed records of all income and expenses</li>
              <li>Consider depreciation deductions for property improvements</li>
              <li>Track mileage for property-related travel</li>
              <li>Save receipts for all deductible expenses</li>
              <li>Consider quarterly estimated tax payments if needed</li>
            </ul>
          </div>
          
          <div style="text-align: center; margin-top: 20px;">
            <button onclick="google.script.run.exportTaxData()" style="margin: 5px; padding: 10px 20px; background: #4CAF50; color: white; border: none; border-radius: 4px;">
              Export Tax Data
            </button>
          </div>
        </div>
      `)
        .setWidth(800)
        .setHeight(700);
      
      SpreadsheetApp.getUi().showModalDialog(html, 'Tax Report');
      
    } catch (error) {
      handleSystemError(error, 'generateTaxReport');
    }
  },
  
  /**
   * Prepare tax data from financial records
   */
  prepareTaxData: function(data, startDate, endDate) {
    const taxData = {
      totalIncome: 0,
      totalDeductions: 0,
      netIncome: 0,
      incomeByCategory: {},
      deductibleExpenses: {}
    };
    
    // Define deductible expense categories
    const deductibleCategories = [
      'Maintenance', 'Utilities', 'Insurance', 'Property Tax',
      'Advertising', 'Professional Services', 'Supplies',
      'Repairs', 'Management Fees', 'Legal Fees'
    ];
    
    data.forEach(row => {
      const date = new Date(row[this.COL.DATE - 1]);
      const amount = row[this.COL.AMOUNT - 1] || 0;
      const category = row[this.COL.CATEGORY - 1] || 'Other';
      
      if (date >= startDate && date <= endDate) {
        if (amount > 0) {
          // Income
          taxData.totalIncome += amount;
          taxData.incomeByCategory[category] = (taxData.incomeByCategory[category] || 0) + amount;
        } else if (amount < 0 && deductibleCategories.includes(category)) {
          // Deductible expense
          const expenseAmount = Math.abs(amount);
          taxData.totalDeductions += expenseAmount;
          taxData.deductibleExpenses[category] = (taxData.deductibleExpenses[category] || 0) + expenseAmount;
        }
      }
    });
    
    taxData.netIncome = taxData.totalIncome - taxData.totalDeductions;
    
    return taxData;
  },
  
  /**
   * Show occupancy analytics
   */
  showOccupancyAnalytics: function() {
    try {
      // Get tenant occupancy data
      const tenantData = SheetManager.getAllData(CONFIG.SHEETS.TENANTS);
      const guestRoomData = SheetManager.getAllData(CONFIG.SHEETS.GUEST_ROOMS);
      const guestBookingData = SheetManager.getAllData(CONFIG.SHEETS.GUEST_BOOKINGS);
      
      const occupancyStats = this.calculateOccupancyStats(tenantData, guestRoomData, guestBookingData);
      
      const html = HtmlService.createHtmlOutput(`
        <div style="font-family: Arial, sans-serif; padding: 20px; line-height: 1.6;">
          <h2>🏠 Occupancy Analytics</h2>
          
          <div style="display: grid; grid-template-columns: repeat(3, 1fr); gap: 20px; margin-bottom: 30px;">
            <div style="background: #e8f5e8; padding: 15px; border-radius: 8px; text-align: center;">
              <h3 style="margin: 0; color: #2e7d32;">Overall Occupancy</h3>
              <p style="font-size: 24px; margin: 5px 0; font-weight: bold;">${occupancyStats.overallOccupancy}%</p>
              <small>${occupancyStats.totalOccupied}/${occupancyStats.totalRooms} rooms</small>
            </div>
            <div style="background: #e3f2fd; padding: 15px; border-radius: 8px; text-align: center;">
              <h3 style="margin: 0; color: #1565c0;">Long-term Rentals</h3>
              <p style="font-size: 24px; margin: 5px 0; font-weight: bold;">${occupancyStats.tenantOccupancy}%</p>
              <small>${occupancyStats.tenantsOccupied}/${occupancyStats.totalTenantRooms} rooms</small>
            </div>
            <div style="background: #fff3e0; padding: 15px; border-radius: 8px; text-align: center;">
              <h3 style="margin: 0; color: #ef6c00;">Guest Rooms</h3>
              <p style="font-size: 24px; margin: 5px 0; font-weight: bold;">${occupancyStats.guestOccupancy}%</p>
              <small>${occupancyStats.guestsOccupied}/${occupancyStats.totalGuestRooms} rooms</small>
            </div>
          </div>
          
          <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 30px;">
            <div>
              <h3>📊 Revenue Efficiency</h3>
              <div style="background: #f5f5f5; padding: 15px; border-radius: 8px;">
                <p><strong>Revenue per Available Room (RevPAR):</strong></p>
                <ul>
                  <li>Long-term: ${Utils.formatCurrency(occupancyStats.tenantRevPAR)}/month</li>
                  <li>Guest rooms: ${Utils.formatCurrency(occupancyStats.guestRevPAR)}/night</li>
                </ul>
                
                <p><strong>Average Daily Rate (ADR):</strong></p>
                <ul>
                  <li>Long-term equivalent: ${Utils.formatCurrency(occupancyStats.tenantADR)}/night</li>
                  <li>Guest rooms: ${Utils.formatCurrency(occupancyStats.guestADR)}/night</li>
                </ul>
              </div>
            </div>
            
            <div>
              <h3>🎯 Performance Targets</h3>
              <div style="background: #f5f5f5; padding: 15px; border-radius: 8px;">
                <div style="margin-bottom: 15px;">
                  <p><strong>Target Overall Occupancy: 85%</strong></p>
                  <div style="background: #ddd; height: 10px; border-radius: 5px;">
                    <div style="background: ${occupancyStats.overallOccupancy >= 85 ? '#4caf50' : '#ff9800'}; height: 100%; width: ${Math.min(occupancyStats.overallOccupancy, 100)}%; border-radius: 5px;"></div>
                  </div>
                  <small>${occupancyStats.overallOccupancy >= 85 ? '✅ Target Met' : `📈 ${(85 - occupancyStats.overallOccupancy).toFixed(1)}% to target`}</small>
                </div>
                
                <div style="margin-bottom: 15px;">
                  <p><strong>Target Guest Room Occupancy: 70%</strong></p>
                  <div style="background: #ddd; height: 10px; border-radius: 5px;">
                    <div style="background: ${occupancyStats.guestOccupancy >= 70 ? '#4caf50' : '#ff9800'}; height: 100%; width: ${Math.min(occupancyStats.guestOccupancy, 100)}%; border-radius: 5px;"></div>
                  </div>
                  <small>${occupancyStats.guestOccupancy >= 70 ? '✅ Target Met' : `📈 ${(70 - occupancyStats.guestOccupancy).toFixed(1)}% to target`}</small>
                </div>
              </div>
            </div>
          </div>
          
          <h3>📈 Optimization Opportunities</h3>
          <div style="background: #e3f2fd; padding: 15px; border-radius: 8px;">
            ${this.generateOccupancyInsights(occupancyStats)}
          </div>
          
          <div style="text-align: center; margin-top: 20px;">
            <button onclick="google.script.run.showProfitabilityDashboard()" style="margin: 5px; padding: 10px 20px;">Profitability Analysis</button>
            <button onclick="google.script.run.analyzeGuestRoomPricing()" style="margin: 5px; padding: 10px 20px;">Pricing Strategy</button>
          </div>
        </div>
      `)
        .setWidth(800)
        .setHeight(600);
      
      SpreadsheetApp.getUi().showModalDialog(html, 'Occupancy Analytics');
      
    } catch (error) {
      handleSystemError(error, 'showOccupancyAnalytics');
    }
  },
  
  /**
   * Calculate occupancy statistics
   */
  calculateOccupancyStats: function(tenantData, guestRoomData, guestBookingData) {
    const stats = {
      totalRooms: 0,
      totalOccupied: 0,
      totalTenantRooms: 0,
      tenantsOccupied: 0,
      totalGuestRooms: 0,
      guestsOccupied: 0,
      overallOccupancy: 0,
      tenantOccupancy: 0,
      guestOccupancy: 0,
      tenantRevPAR: 0,
      guestRevPAR: 0,
      tenantADR: 0,
      guestADR: 0
    };
    
    // Calculate tenant room statistics
    let tenantRevenue = 0;
    tenantData.forEach(tenant => {
      if (tenant[0]) { // Has room number
        stats.totalTenantRooms++;
        const roomStatus = tenant[8];
        const rent = tenant[2] || tenant[1] || 0; // Negotiated or standard rent
        
        if (roomStatus === CONFIG.STATUS.ROOM.OCCUPIED) {
          stats.tenantsOccupied++;
          tenantRevenue += rent;
        }
      }
    });
    
    // Calculate guest room statistics
    let guestRevenue = 0;
    let guestNights = 0;
    
    guestRoomData.forEach(room => {
      if (room[1]) { // Has room number
        stats.totalGuestRooms++;
        if (room[9] === 'Occupied') {
          stats.guestsOccupied++;
        }
      }
    });
    
    // Calculate guest room revenue from recent bookings
    const now = new Date();
    const monthStart = new Date(now.getFullYear(), now.getMonth(), 1);
    
    guestBookingData.forEach(booking => {
      const checkIn = new Date(booking[6]);
      const amount = booking[12] || 0;
      const nights = booking[8] || 0;
      
      if (checkIn >= monthStart && amount > 0) {
        guestRevenue += amount;
        guestNights += nights;
      }
    });
    
    // Calculate totals
    stats.totalRooms = stats.totalTenantRooms + stats.totalGuestRooms;
    stats.totalOccupied = stats.tenantsOccupied + stats.guestsOccupied;
    
    // Calculate occupancy rates
    stats.overallOccupancy = stats.totalRooms > 0 ? 
      Math.round((stats.totalOccupied / stats.totalRooms) * 100) : 0;
    stats.tenantOccupancy = stats.totalTenantRooms > 0 ? 
      Math.round((stats.tenantsOccupied / stats.totalTenantRooms) * 100) : 0;
    stats.guestOccupancy = stats.totalGuestRooms > 0 ? 
      Math.round((stats.guestsOccupied / stats.totalGuestRooms) * 100) : 0;
    
    // Calculate revenue metrics
    stats.tenantRevPAR = stats.totalTenantRooms > 0 ? tenantRevenue / stats.totalTenantRooms : 0;
    stats.guestRevPAR = stats.totalGuestRooms > 0 ? guestRevenue / stats.totalGuestRooms : 0;
    stats.tenantADR = stats.tenantsOccupied > 0 ? tenantRevenue / stats.tenantsOccupied / 30 : 0; // Daily equivalent
    stats.guestADR = guestNights > 0 ? guestRevenue / guestNights : 0;
    
    return stats;
  },
  
  /**
   * Generate occupancy insights and recommendations
   */
  generateOccupancyInsights: function(stats) {
    const insights = [];
    
    // Overall occupancy insights
    if (stats.overallOccupancy < 70) {
      insights.push('🔴 <strong>Low overall occupancy.</strong> Focus on marketing and competitive pricing.');
    } else if (stats.overallOccupancy > 90) {
      insights.push('🟢 <strong>Excellent occupancy!</strong> Consider raising rates or expanding capacity.');
    }
    
    // Tenant vs guest room balance
    if (stats.tenantOccupancy > 95 && stats.guestOccupancy < 50) {
      insights.push('💡 <strong>Consider converting guest rooms</strong> to long-term rentals given high demand.');
    } else if (stats.guestOccupancy > 80 && stats.tenantOccupancy < 70) {
      insights.push('💡 <strong>Guest rooms performing well.</strong> Consider expanding short-term offerings.');
    }
    
    // Revenue optimization
    if (stats.guestADR > stats.tenantADR * 1.5) {
      insights.push('💰 <strong>Guest rooms generate higher daily rates.</strong> Optimize guest room mix.');
    }
    
    // Performance benchmarks
    if (stats.tenantRevPAR < 500) {
      insights.push('📈 <strong>Long-term revenue below benchmark.</strong> Review rent levels and vacancy reduction.');
    }
    
    if (insights.length === 0) {
      insights.push('✅ <strong>Occupancy performance is balanced.</strong> Continue monitoring for optimization opportunities.');
    }
    
    return '<ul><li>' + insights.join('</li><li>') + '</li></ul>';
  },
  
  /**
   * Show profitability dashboard
   */
  showProfitabilityDashboard: function() {
    try {
      const data = SheetManager.getAllData(CONFIG.SHEETS.BUDGET);
      const now = new Date();
      
      // Calculate profitability metrics for different time periods
      const monthStart = new Date(now.getFullYear(), now.getMonth(), 1);
      const quarterStart = new Date(now.getFullYear(), Math.floor(now.getMonth() / 3) * 3, 1);
      const yearStart = new Date(now.getFullYear(), 0, 1);
      
      const monthlyMetrics = this.calculateProfitabilityMetrics(data, monthStart, now);
      const quarterlyMetrics = this.calculateProfitabilityMetrics(data, quarterStart, now);
      const yearlyMetrics = this.calculateProfitabilityMetrics(data, yearStart, now);
      
      const html = HtmlService.createHtmlOutput(`
        <div style="font-family: Arial, sans-serif; padding: 20px; line-height: 1.6;">
          <h2>💡 Profitability Dashboard</h2>
          
          <div style="display: grid; grid-template-columns: repeat(3, 1fr); gap: 20px; margin-bottom: 30px;">
            <div style="background: #e8f5e8; padding: 15px; border-radius: 8px; text-align: center;">
              <h4 style="margin: 0; color: #2e7d32;">This Month</h4>
              <p style="font-size: 20px; margin: 5px 0; font-weight: bold;">${Utils.formatCurrency(monthlyMetrics.netProfit)}</p>
              <small>Margin: ${monthlyMetrics.profitMargin}%</small>
            </div>
            <div style="background: #fff3e0; padding: 15px; border-radius: 8px; text-align: center;">
              <h4 style="margin: 0; color: #ef6c00;">This Quarter</h4>
              <p style="font-size: 20px; margin: 5px 0; font-weight: bold;">${Utils.formatCurrency(quarterlyMetrics.netProfit)}</p>
              <small>Margin: ${quarterlyMetrics.profitMargin}%</small>
            </div>
            <div style="background: #e3f2fd; padding: 15px; border-radius: 8px; text-align: center;">
              <h4 style="margin: 0; color: #1565c0;">This Year</h4>
              <p style="font-size: 20px; margin: 5px 0; font-weight: bold;">${Utils.formatCurrency(yearlyMetrics.netProfit)}</p>
              <small>Margin: ${yearlyMetrics.profitMargin}%</small>
            </div>
          </div>
          
          <div style="background: #f5f5f5; padding: 15px; border-radius: 8px; margin-bottom: 20px;">
            <h3>📊 Key Performance Metrics (YTD)</h3>
            <table style="width: 100%; border-collapse: collapse;">
              <tr style="border-bottom: 1px solid #ddd;">
                <th style="text-align: left; padding: 8px;">Metric</th>
                <th style="text-align: right; padding: 8px;">Value</th>
                <th style="text-align: center; padding: 8px;">Status</th>
              </tr>
              <tr>
                <td style="padding: 8px;">Gross Revenue</td>
                <td style="text-align: right; padding: 8px;">${Utils.formatCurrency(yearlyMetrics.totalRevenue)}</td>
                <td style="text-align: center; padding: 8px;">📈</td>
              </tr>
              <tr>
                <td style="padding: 8px;">Operating Expenses</td>
                <td style="text-align: right; padding: 8px;">${Utils.formatCurrency(yearlyMetrics.totalExpenses)}</td>
                <td style="text-align: center; padding: 8px;">💸</td>
              </tr>
              <tr style="border-top: 2px solid #ddd; font-weight: bold;">
                <td style="padding: 8px;">Net Operating Income</td>
                <td style="text-align: right; padding: 8px; color: ${yearlyMetrics.netProfit >= 0 ? '#2e7d32' : '#c62828'};">${Utils.formatCurrency(yearlyMetrics.netProfit)}</td>
                <td style="text-align: center; padding: 8px;">${yearlyMetrics.netProfit >= 0 ? '✅' : '⚠️'}</td>
              </tr>
              <tr>
                <td style="padding: 8px;">Profit Margin</td>
                <td style="text-align: right; padding: 8px; color: ${yearlyMetrics.profitMargin >= 20 ? '#2e7d32' : yearlyMetrics.profitMargin >= 10 ? '#ef6c00' : '#c62828'};">${yearlyMetrics.profitMargin}%</td>
                <td style="text-align: center; padding: 8px;">${yearlyMetrics.profitMargin >= 20 ? '🌟' : yearlyMetrics.profitMargin >= 10 ? '👍' : '⚠️'}</td>
              </tr>
              <tr>
                <td style="padding: 8px;">Cash Flow Ratio</td>
                <td style="text-align: right; padding: 8px;">${yearlyMetrics.cashFlowRatio.toFixed(2)}x</td>
                <td style="text-align: center; padding: 8px;">${yearlyMetrics.cashFlowRatio >= 1.2 ? '💪' : yearlyMetrics.cashFlowRatio >= 1.0 ? '👌' : '🚨'}</td>
              </tr>
            </table>
          </div>
          
          <div style="background: #e3f2fd; padding: 15px; border-radius: 8px; margin-bottom: 20px;">
            <h4>💡 Profitability Insights</h4>
            ${this.generateProfitabilityInsights(monthlyMetrics, quarterlyMetrics, yearlyMetrics)}
          </div>
          
          <div style="text-align: center;">
            <button onclick="google.script.run.generateMonthlyFinancialReport()" style="margin: 5px; padding: 10px 20px;">Detailed Financial Report</button>
            <button onclick="google.script.run.exportFinancialData()" style="margin: 5px; padding: 10px 20px;">Export Financial Data</button>
          </div>
        </div>
      `)
        .setWidth(900)
        .setHeight(700);
      
      SpreadsheetApp.getUi().showModalDialog(html, 'Profitability Dashboard');
      
    } catch (error) {
      handleSystemError(error, 'showProfitabilityDashboard');
    }
  },
  
  /**
   * Calculate profitability metrics for a given period
   */
  calculateProfitabilityMetrics: function(data, startDate, endDate) {
    let totalRevenue = 0;
    let totalExpenses = 0;
    
    data.forEach(row => {
      const date = new Date(row[this.COL.DATE - 1]);
      const amount = row[this.COL.AMOUNT - 1] || 0;
      
      if (date >= startDate && date <= endDate) {
        if (amount > 0) {
          totalRevenue += amount;
        } else {
          totalExpenses += Math.abs(amount);
        }
      }
    });
    
    const netProfit = totalRevenue - totalExpenses;
    const profitMargin = totalRevenue > 0 ? Math.round((netProfit / totalRevenue) * 100) : 0;
    const cashFlowRatio = totalExpenses > 0 ? totalRevenue / totalExpenses : 0;
    
    // Estimate ROI (simplified calculation)
    const estimatedPropertyValue = totalRevenue * 12 * 10; // 10x annual revenue estimate
    const roi = estimatedPropertyValue > 0 ? Math.round((netProfit * 12 / estimatedPropertyValue) * 100) : 0;
    
    // Calculate break-even point
    const monthsInPeriod = Math.max(1, Math.ceil((endDate - startDate) / (30 * 24 * 60 * 60 * 1000)));
    const breakEven = totalExpenses / monthsInPeriod;
    
    return {
      totalRevenue,
      totalExpenses,
      netProfit,
      profitMargin,
      cashFlowRatio,
      roi,
      breakEven
    };
  },
  
  /**
   * Generate profitability insights
   */
  generateProfitabilityInsights: function(monthly, quarterly, yearly) {
    const insights = [];
    
    // Profit margin analysis
    if (yearly.profitMargin > 25) {
      insights.push('🌟 <strong>Excellent profit margins!</strong> Your property is highly profitable.');
    } else if (yearly.profitMargin < 10) {
      insights.push('⚠️ <strong>Low profit margins.</strong> Review expenses and consider rent increases.');
    }
    
    // Growth trend analysis
    const monthlyAverage = yearly.netProfit / Math.max(1, new Date().getMonth() + 1);
    if (monthly.netProfit > monthlyAverage * 1.2) {
      insights.push('📈 <strong>Strong performance this month!</strong> Above average profitability.');
    } else if (monthly.netProfit < monthlyAverage * 0.8) {
      insights.push('📉 <strong>Below-average month.</strong> Review recent changes and seasonal factors.');
    }
    
    // Cash flow analysis
    if (yearly.cashFlowRatio < 1.0) {
      insights.push('🚨 <strong>Negative cash flow.</strong> Immediate action needed to reduce expenses or increase revenue.');
    } else if (yearly.cashFlowRatio > 1.5) {
      insights.push('💰 <strong>Strong cash flow.</strong> Consider reinvestment opportunities.');
    }
    
    // ROI analysis
    if (yearly.roi > 12) {
      insights.push('🎯 <strong>Excellent ROI!</strong> Your investment is performing very well.');
    } else if (yearly.roi < 6) {
      insights.push('📊 <strong>ROI below market average.</strong> Consider optimization strategies.');
    }
    
    if (insights.length === 0) {
      insights.push('📊 <strong>Stable profitability.</strong> Continue monitoring for optimization opportunities.');
    }
    
    return '<ul><li>' + insights.join('</li><li>') + '</li></ul>';
  },

  /**
   * Gather all dashboard data at once for performance
   */
  gatherDashboardData: function() {
    const now = new Date();
    const data = SheetManager.getAllData(CONFIG.SHEETS.BUDGET);

    const monthStart = new Date(now.getFullYear(), now.getMonth(), 1);
    const prevMonthStart = new Date(now.getFullYear(), now.getMonth() - 1, 1);
    const prevMonthEnd = new Date(now.getFullYear(), now.getMonth(), 0);
    const quarterStart = new Date(now.getFullYear(), Math.floor(now.getMonth() / 3) * 3, 1);
    const yearStart = new Date(now.getFullYear(), 0, 1);

    const tenantData = SheetManager.getAllData(CONFIG.SHEETS.TENANTS);
    const guestRoomData = SheetManager.getAllData(CONFIG.SHEETS.GUEST_ROOMS);
    const guestBookingData = SheetManager.getAllData(CONFIG.SHEETS.GUEST_BOOKINGS);

    return {
      monthYear: Utils.formatDate(now, 'MMMM yyyy'),
      monthAnalysis: this.analyzeFinancialData(data, monthStart, now),
      previousMonth: this.analyzeFinancialData(data, prevMonthStart, prevMonthEnd),
      ytdAnalysis: this.analyzeFinancialData(data, yearStart, now),
      quarterAnalysis: this.analyzeFinancialData(data, quarterStart, now),
      occupancyStats: this.calculateOccupancyStats(tenantData, guestRoomData, guestBookingData),
      monthlyProfit: this.calculateProfitabilityMetrics(data, monthStart, now),
      quarterlyProfit: this.calculateProfitabilityMetrics(data, quarterStart, now),
      yearlyProfit: this.calculateProfitabilityMetrics(data, yearStart, now)
    };
  },

  /**
   * Build monthly report section HTML
   */
  buildMonthlySection: function(d) {
    const r = d.monthAnalysis;
    return `
      <h3>💰 Monthly Financial Report - ${d.monthYear}</h3>
      <div style="display: grid; grid-template-columns: repeat(3, 1fr); gap: 20px; margin-bottom: 30px;">
        <div style="background: #e8f5e8; padding: 15px; border-radius: 8px; text-align: center;">
          <h4 style="margin: 0; color: #2e7d32;">Total Income</h4>
          <p style="font-size: 24px; margin: 5px 0; font-weight: bold;">${Utils.formatCurrency(r.totalIncome)}</p>
          <small>${r.incomeTransactions} transactions</small>
        </div>
        <div style="background: #ffebee; padding: 15px; border-radius: 8px; text-align: center;">
          <h4 style="margin: 0; color: #c62828;">Total Expenses</h4>
          <p style="font-size: 24px; margin: 5px 0; font-weight: bold;">${Utils.formatCurrency(r.totalExpenses)}</p>
          <small>${r.expenseTransactions} transactions</small>
        </div>
        <div style="background: #e3f2fd; padding: 15px; border-radius: 8px; text-align: center;">
          <h4 style="margin: 0; color: #1565c0;">Net Profit</h4>
          <p style="font-size: 24px; margin: 5px 0; font-weight: bold; color: ${r.netProfit >= 0 ? '#2e7d32' : '#c62828'};">${Utils.formatCurrency(r.netProfit)}</p>
          <small>Profit Margin: ${r.profitMargin}%</small>
        </div>
      </div>

      <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 30px;">
        <div>
          <h4>📈 Income Breakdown</h4>
          <div style="background: #f5f5f5; padding: 15px; border-radius: 8px;">
            ${Object.entries(r.incomeByCategory).map(([c,a]) => `<div style="display:flex;justify-content:space-between;margin-bottom:8px;"><span><strong>${c}:</strong></span><span>${Utils.formatCurrency(a)}</span></div>`).join('')}
          </div>
        </div>
        <div>
          <h4>📉 Expense Breakdown</h4>
          <div style="background: #f5f5f5; padding: 15px; border-radius: 8px;">
            ${Object.entries(r.expensesByCategory).map(([c,a]) => `<div style="display:flex;justify-content:space-between;margin-bottom:8px;"><span><strong>${c}:</strong></span><span>${Utils.formatCurrency(a)}</span></div>`).join('')}
          </div>
        </div>
      </div>
    `;
  },

  /**
   * Build revenue analysis section HTML
   */
  buildRevenueSection: function(d) {
    const c = d.monthAnalysis;
    const p = d.previousMonth;
    const y = d.ytdAnalysis;
    const incomeChange = p.totalIncome > 0 ? ((c.totalIncome - p.totalIncome) / p.totalIncome * 100) : 0;
    return `
      <h3>📈 Revenue Analysis</h3>
      <div style="display: grid; grid-template-columns: repeat(3, 1fr); gap: 20px; margin-bottom: 30px;">
        <div style="background: #e8f5e8; padding: 15px; border-radius: 8px; text-align: center;">
          <h4 style="margin:0;color:#2e7d32;">Current Month</h4>
          <p style="font-size:20px;margin:5px 0;font-weight:bold;">${Utils.formatCurrency(c.totalIncome)}</p>
          <small style="color:${incomeChange >= 0 ? '#2e7d32' : '#c62828'};">${incomeChange >= 0 ? '↗' : '↘'} ${Math.abs(incomeChange).toFixed(1)}% vs last month</small>
        </div>
        <div style="background: #fff3e0; padding: 15px; border-radius: 8px; text-align: center;">
          <h4 style="margin:0;color:#ef6c00;">Previous Month</h4>
          <p style="font-size:20px;margin:5px 0;font-weight:bold;">${Utils.formatCurrency(p.totalIncome)}</p>
          <small>Net: ${Utils.formatCurrency(p.netProfit)}</small>
        </div>
        <div style="background: #e3f2fd; padding: 15px; border-radius: 8px; text-align: center;">
          <h4 style="margin:0;color:#1565c0;">Year to Date</h4>
          <p style="font-size:20px;margin:5px 0;font-weight:bold;">${Utils.formatCurrency(y.totalIncome)}</p>
          <small>Avg Monthly: ${Utils.formatCurrency(y.totalIncome / (new Date().getMonth() + 1))}</small>
        </div>
      </div>

      <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 30px;">
        <div>
          <h4>📊 Revenue Streams (YTD)</h4>
          <div style="background:#f5f5f5;padding:15px;border-radius:8px;">
            ${Object.entries(y.incomeByCategory).sort(([,a],[,b])=>b-a).map(([cat,amt]) => {
              const pct = (amt / y.totalIncome * 100).toFixed(1);
              return `<div style="margin-bottom:10px;"><div style="display:flex;justify-content:space-between;margin-bottom:2px;"><span><strong>${cat}</strong></span><span>${Utils.formatCurrency(amt)} (${pct}%)</span></div><div style="background:#ddd;height:8px;border-radius:4px;"><div style="background:#4caf50;height:100%;width:${pct}%;border-radius:4px;"></div></div></div>`;
            }).join('')}
          </div>
        </div>
        <div>
          <h4>💸 Expense Categories (YTD)</h4>
          <div style="background:#f5f5f5;padding:15px;border-radius:8px;">
            ${Object.entries(y.expensesByCategory).sort(([,a],[,b])=>b-a).map(([cat,amt]) => {
              const pct = (amt / y.totalExpenses * 100).toFixed(1);
              return `<div style="margin-bottom:10px;"><div style="display:flex;justify-content:space-between;margin-bottom:2px;"><span><strong>${cat}</strong></span><span>${Utils.formatCurrency(amt)} (${pct}%)</span></div><div style="background:#ddd;height:8px;border-radius:4px;"><div style="background:#f44336;height:100%;width:${pct}%;border-radius:4px;"></div></div></div>`;
            }).join('')}
          </div>
        </div>
      </div>
    `;
  },

  /**
   * Build occupancy metrics section HTML
   */
  buildOccupancySection: function(d) {
    const o = d.occupancyStats;
    return `
      <h3>🏠 Occupancy Analytics</h3>
      <div style="display: grid; grid-template-columns: repeat(3, 1fr); gap: 20px; margin-bottom: 30px;">
        <div style="background:#e8f5e8;padding:15px;border-radius:8px;text-align:center;">
          <h4 style="margin:0;color:#2e7d32;">Overall Occupancy</h4>
          <p style="font-size:24px;margin:5px 0;font-weight:bold;">${o.overallOccupancy}%</p>
          <small>${o.totalOccupied}/${o.totalRooms} rooms</small>
        </div>
        <div style="background:#e3f2fd;padding:15px;border-radius:8px;text-align:center;">
          <h4 style="margin:0;color:#1565c0;">Long-term Rentals</h4>
          <p style="font-size:24px;margin:5px 0;font-weight:bold;">${o.tenantOccupancy}%</p>
          <small>${o.tenantsOccupied}/${o.totalTenantRooms} rooms</small>
        </div>
        <div style="background:#fff3e0;padding:15px;border-radius:8px;text-align:center;">
          <h4 style="margin:0;color:#ef6c00;">Guest Rooms</h4>
          <p style="font-size:24px;margin:5px 0;font-weight:bold;">${o.guestOccupancy}%</p>
          <small>${o.guestsOccupied}/${o.totalGuestRooms} rooms</small>
        </div>
      </div>

      <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 30px;">
        <div>
          <h4>📊 Revenue Efficiency</h4>
          <div style="background:#f5f5f5;padding:15px;border-radius:8px;">
            <p><strong>Revenue per Available Room (RevPAR):</strong></p>
            <ul>
              <li>Long-term: ${Utils.formatCurrency(o.tenantRevPAR)}/month</li>
              <li>Guest rooms: ${Utils.formatCurrency(o.guestRevPAR)}/night</li>
            </ul>
            <p><strong>Average Daily Rate (ADR):</strong></p>
            <ul>
              <li>Long-term equivalent: ${Utils.formatCurrency(o.tenantADR)}/night</li>
              <li>Guest rooms: ${Utils.formatCurrency(o.guestADR)}/night</li>
            </ul>
          </div>
        </div>
        <div>
          <h4>🎯 Performance Targets</h4>
          <div style="background:#f5f5f5;padding:15px;border-radius:8px;">
            <div style="margin-bottom:15px;">
              <p><strong>Target Overall Occupancy: 85%</strong></p>
              <div style="background:#ddd;height:10px;border-radius:5px;">
                <div style="background:${o.overallOccupancy >= 85 ? '#4caf50' : '#ff9800'};height:100%;width:${Math.min(o.overallOccupancy,100)}%;border-radius:5px;"></div>
              </div>
              <small>${o.overallOccupancy >= 85 ? '✅ Target Met' : `📈 ${(85 - o.overallOccupancy).toFixed(1)}% to target`}</small>
            </div>
            <div style="margin-bottom:15px;">
              <p><strong>Target Guest Room Occupancy: 70%</strong></p>
              <div style="background:#ddd;height:10px;border-radius:5px;">
                <div style="background:${o.guestOccupancy >= 70 ? '#4caf50' : '#ff9800'};height:100%;width:${Math.min(o.guestOccupancy,100)}%;border-radius:5px;"></div>
              </div>
              <small>${o.guestOccupancy >= 70 ? '✅ Target Met' : `📈 ${(70 - o.guestOccupancy).toFixed(1)}% to target`}</small>
            </div>
          </div>
        </div>
      </div>

      <h4>📈 Optimization Opportunities</h4>
      <div style="background:#e3f2fd;padding:15px;border-radius:8px;">
        ${this.generateOccupancyInsights(o)}
      </div>
    `;
  },

  /**
   * Build profitability section HTML
   */
  buildProfitSection: function(d) {
    const m = d.monthlyProfit;
    const q = d.quarterlyProfit;
    const y = d.yearlyProfit;
    return `
      <h3>💡 Profitability Dashboard</h3>
      <div style="display: grid; grid-template-columns: repeat(3, 1fr); gap: 20px; margin-bottom: 30px;">
        <div style="background:#e8f5e8;padding:15px;border-radius:8px;text-align:center;">
          <h4 style="margin:0;color:#2e7d32;">This Month</h4>
          <p style="font-size:20px;margin:5px 0;font-weight:bold;">${Utils.formatCurrency(m.netProfit)}</p>
          <small>Margin: ${m.profitMargin}%</small>
        </div>
        <div style="background:#fff3e0;padding:15px;border-radius:8px;text-align:center;">
          <h4 style="margin:0;color:#ef6c00;">This Quarter</h4>
          <p style="font-size:20px;margin:5px 0;font-weight:bold;">${Utils.formatCurrency(q.netProfit)}</p>
          <small>Margin: ${q.profitMargin}%</small>
        </div>
        <div style="background:#e3f2fd;padding:15px;border-radius:8px;text-align:center;">
          <h4 style="margin:0;color:#1565c0;">This Year</h4>
          <p style="font-size:20px;margin:5px 0;font-weight:bold;">${Utils.formatCurrency(y.netProfit)}</p>
          <small>Margin: ${y.profitMargin}%</small>
        </div>
      </div>

      <div style="background:#f5f5f5;padding:15px;border-radius:8px;margin-bottom:20px;">
        <h4>📊 Key Performance Metrics (YTD)</h4>
        <table style="width:100%;border-collapse:collapse;">
          <tr style="border-bottom:1px solid #ddd;">
            <th style="text-align:left;padding:8px;">Metric</th>
            <th style="text-align:right;padding:8px;">Value</th>
            <th style="text-align:center;padding:8px;">Status</th>
          </tr>
          <tr>
            <td style="padding:8px;">Gross Revenue</td>
            <td style="text-align:right;padding:8px;">${Utils.formatCurrency(y.totalRevenue)}</td>
            <td style="text-align:center;padding:8px;">📈</td>
          </tr>
          <tr>
            <td style="padding:8px;">Operating Expenses</td>
            <td style="text-align:right;padding:8px;">${Utils.formatCurrency(y.totalExpenses)}</td>
            <td style="text-align:center;padding:8px;">💸</td>
          </tr>
          <tr style="border-top:2px solid #ddd;font-weight:bold;">
            <td style="padding:8px;">Net Operating Income</td>
            <td style="text-align:right;padding:8px;color:${y.netProfit >= 0 ? '#2e7d32' : '#c62828'};">${Utils.formatCurrency(y.netProfit)}</td>
            <td style="text-align:center;padding:8px;">${y.netProfit >= 0 ? '✅' : '⚠️'}</td>
          </tr>
          <tr>
            <td style="padding:8px;">Profit Margin</td>
            <td style="text-align:right;padding:8px;color:${y.profitMargin >= 20 ? '#2e7d32' : y.profitMargin >= 10 ? '#ef6c00' : '#c62828'};">${y.profitMargin}%</td>
            <td style="text-align:center;padding:8px;">${y.profitMargin >= 20 ? '🌟' : y.profitMargin >= 10 ? '👍' : '⚠️'}</td>
          </tr>
          <tr>
            <td style="padding:8px;">Cash Flow Ratio</td>
            <td style="text-align:right;padding:8px;">${y.cashFlowRatio.toFixed(2)}x</td>
            <td style="text-align:center;padding:8px;">${y.cashFlowRatio >= 1.2 ? '💪' : y.cashFlowRatio >= 1.0 ? '👌' : '🚨'}</td>
          </tr>
        </table>
      </div>

      <div style="background:#e3f2fd;padding:15px;border-radius:8px;margin-bottom:20px;">
        <h4>💡 Profitability Insights</h4>
        ${this.generateProfitabilityInsights(m, q, y)}
      </div>
    `;
  },
  
  /**
   * Export financial data for external analysis
   */
  exportFinancialData: function() {
    try {
      const ui = SpreadsheetApp.getUi();
      
      const response = ui.alert(
        'Export Financial Data',
        'This will create a CSV file with all financial transactions. Continue?',
        ui.ButtonSet.YES_NO
      );
      
      if (response !== ui.Button.YES) return;
      
      const data = SheetManager.getAllData(CONFIG.SHEETS.BUDGET, true); // Include headers
      
      // Convert data to CSV format
      const csvContent = data.map(row => {
        return row.map(cell => {
          // Handle dates, numbers, and text
          if (cell instanceof Date) {
            return Utils.formatDate(cell, 'yyyy-MM-dd');
          } else if (typeof cell === 'string' && cell.includes(',')) {
            return `"${cell}"`;
          }
          return cell;
        }).join(',');
      }).join('\n');
      
      // Create file
      const fileName = `FinancialData_${Utils.formatDate(new Date(), 'yyyyMMdd_HHmmss')}.csv`;
      const blob = Utilities.newBlob(csvContent, 'text/csv', fileName);
      const file = DriveApp.createFile(blob);
      
      ui.alert(
        'Export Complete',
        `Financial data exported to: ${fileName}\n\nFile ID: ${file.getId()}\n\nYou can find this file in your Google Drive.`,
        ui.ButtonSet.OK
      );
      
    } catch (error) {
      handleSystemError(error, 'exportFinancialData');
    }
  },
  
  /**
   * Export tax data for accountant
   */
  exportTaxData: function() {
    try {
      const ui = SpreadsheetApp.getUi();
      const now = new Date();
      const taxYear = now.getFullYear();
      
      const data = SheetManager.getAllData(CONFIG.SHEETS.BUDGET);
      const yearStart = new Date(taxYear, 0, 1);
      const yearEnd = new Date(taxYear, 11, 31);
      
      const taxData = this.prepareTaxData(data, yearStart, yearEnd);
      
      // Create comprehensive tax report
      const taxReport = {
        reportInfo: {
          propertyName: CONFIG.SYSTEM.PROPERTY_NAME,
          taxYear: taxYear,
          generatedDate: new Date().toISOString(),
          reportType: 'Schedule E - Rental Property'
        },
        summary: {
          totalRentalIncome: taxData.totalIncome,
          totalDeductibleExpenses: taxData.totalDeductions,
          netRentalIncome: taxData.netIncome
        },
        incomeDetail: taxData.incomeByCategory,
        expenseDetail: taxData.deductibleExpenses,
        transactions: []
      };
      
      // Include detailed transactions for tax year
      data.forEach(row => {
        const date = new Date(row[this.COL.DATE - 1]);
        if (date >= yearStart && date <= yearEnd) {
          taxReport.transactions.push({
            date: Utils.formatDate(date, 'yyyy-MM-dd'),
            type: row[this.COL.TYPE - 1],
            description: row[this.COL.DESCRIPTION - 1],
            amount: row[this.COL.AMOUNT - 1],
            category: row[this.COL.CATEGORY - 1],
            paymentMethod: row[this.COL.PAYMENT_METHOD - 1],
            reference: row[this.COL.REFERENCE_NUMBER - 1]
          });
        }
      });
      
      // Create JSON file for detailed data
      const jsonFileName = `TaxData_${taxYear}_${Utils.formatDate(new Date(), 'yyyyMMdd')}.json`;
      const jsonBlob = Utilities.newBlob(JSON.stringify(taxReport, null, 2), 'application/json', jsonFileName);
      const jsonFile = DriveApp.createFile(jsonBlob);
      
      // Create CSV file for spreadsheet compatibility
      const csvData = [
        ['Date', 'Type', 'Description', 'Amount', 'Category', 'Payment Method', 'Reference'],
        ...taxReport.transactions.map(t => [t.date, t.type, t.description, t.amount, t.category, t.paymentMethod, t.reference])
      ];
      
      const csvContent = csvData.map(row => 
        row.map(cell => typeof cell === 'string' && cell.includes(',') ? `"${cell}"` : cell).join(',')
      ).join('\n');
      
      const csvFileName = `TaxTransactions_${taxYear}_${Utils.formatDate(new Date(), 'yyyyMMdd')}.csv`;
      const csvBlob = Utilities.newBlob(csvContent, 'text/csv', csvFileName);
      const csvFile = DriveApp.createFile(csvBlob);
      
      ui.alert(
        'Tax Data Export Complete',
        `Tax data for ${taxYear} has been exported:\n\n` +
        `Detailed Report: ${jsonFileName}\n` +
        `Transactions CSV: ${csvFileName}\n\n` +
        `Both files are available in your Google Drive.\n` +
        `Share these with your tax professional.`,
        ui.ButtonSet.OK
      );
      
    } catch (error) {
      handleSystemError(error, 'exportTaxData');
    }
  },

  /**
   * Show consolidated financial dashboard
   */
  showFinancialDashboard: function() {
    try {
      const dash = this.gatherDashboardData();

      const html = HtmlService.createHtmlOutput(`
        <style>
          .fd-tab{display:none;}
          .fd-tab.active{display:block;}
          .fd-tabs button{margin-right:5px;}
        </style>
        <div style="font-family: Arial, sans-serif; padding:20px; line-height:1.6;">
          <h2>📊 Financial Dashboard</h2>
          <div class="fd-tabs" style="margin-bottom:15px;">
            <button onclick="fdShow('monthly')">Monthly Report</button>
            <button onclick="fdShow('revenue')">Revenue Analysis</button>
            <button onclick="fdShow('occupancy')">Occupancy Metrics</button>
            <button onclick="fdShow('profit')">Profitability</button>
          </div>

          <div id="monthly" class="fd-tab active">
            ${this.buildMonthlySection(dash)}
          </div>

          <div id="revenue" class="fd-tab">
            ${this.buildRevenueSection(dash)}
          </div>

          <div id="occupancy" class="fd-tab">
            ${this.buildOccupancySection(dash)}
          </div>

          <div id="profit" class="fd-tab">
            ${this.buildProfitSection(dash)}
          </div>

          <div style="text-align:center; margin-top:20px;">
            <button onclick="google.script.run.showExportOptions()" style="padding:8px 20px;">Export / Download</button>
          </div>
        </div>
        <script>
          function fdShow(id){
            document.querySelectorAll('.fd-tab').forEach(t=>t.classList.remove('active'));
            document.getElementById(id).classList.add('active');
          }
        </script>
      `).setWidth(1000).setHeight(800);

      SpreadsheetApp.getUi().showModalDialog(html, 'Financial Dashboard');
    } catch (error) {
      handleSystemError(error, 'showFinancialDashboard');
    }
  },

  /**
   * Show unified export options
   */
  showExportOptions: function() {
    const html = HtmlService.createHtmlOutput(`
      <div style="font-family: Arial, sans-serif; padding:20px;">
        <h3>Export Financial Data</h3>
        <p>
          <button onclick="google.script.run.exportFinancialData()">Full CSV</button>
          <button onclick="google.script.run.exportTaxData()">Current Tax Year</button>
        </p>
        <p>Custom Range:</p>
        <p>
          <input type="date" id="start"> to
          <input type="date" id="end">
          <button onclick="exportRange()">Export</button>
        </p>
        <script>
          function exportRange(){
            const s=document.getElementById('start').value;
            const e=document.getElementById('end').value;
            google.script.run.exportFinancialDataRange(s,e);
          }
        </script>
      </div>
    `).setWidth(400).setHeight(220);
    SpreadsheetApp.getUi().showModalDialog(html, 'Export Options');
  },

  /**
   * Export financial data within a custom range
   */
  exportFinancialDataRange: function(start, end) {
    try {
      const startDate = new Date(start);
      const endDate = new Date(end);

      const data = SheetManager.getAllData(CONFIG.SHEETS.BUDGET, true);
      const filtered = [data[0]];
      for (let i=1; i<data.length; i++) {
        const row = data[i];
        const d = new Date(row[this.COL.DATE - 1]);
        if (d >= startDate && d <= endDate) filtered.push(row);
      }

      const csvContent = filtered.map(row => row.map(cell => {
        if (cell instanceof Date) return Utils.formatDate(cell, 'yyyy-MM-dd');
        if (typeof cell === 'string' && cell.includes(',')) return `"${cell}"`;
        return cell;
      }).join(',')).join('\n');

      const fileName = `FinancialData_${Utils.formatDate(startDate,'yyyyMMdd')}_${Utils.formatDate(endDate,'yyyyMMdd')}.csv`;
      const blob = Utilities.newBlob(csvContent, 'text/csv', fileName);
      DriveApp.createFile(blob);

      SpreadsheetApp.getUi().alert('Export complete: '+fileName);
    } catch (error) {
      handleSystemError(error, 'exportFinancialDataRange');
    }
  }
};

Logger.log('FinancialManager module loaded successfully');
