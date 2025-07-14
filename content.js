// Enhanced Content script for ETA Invoice Exporter - Complete Fix for Bulk Loading
class ETAContentScript {
  constructor() {
    this.invoiceData = [];
    this.allPagesData = [];
    this.totalCount = 0;
    this.currentPage = 1;
    this.totalPages = 1;
    this.resultsPerPage = 50;
    this.isProcessingAllPages = false;
    this.progressCallback = null;
    this.domObserver = null;
    this.pageLoadTimeout = 15000;
    this.init();
  }
  
  init() {
    console.log('ETA Exporter: Content script initialized');
    if (document.readyState === 'loading') {
      document.addEventListener('DOMContentLoaded', () => this.scanForInvoices());
    } else {
      setTimeout(() => this.scanForInvoices(), 1000);
    }
    
    this.setupMutationObserver();
  }
  
  setupMutationObserver() {
    this.observer = new MutationObserver((mutations) => {
      let shouldRescan = false;
      
      mutations.forEach((mutation) => {
        if (mutation.type === 'childList' && mutation.addedNodes.length > 0) {
          mutation.addedNodes.forEach((node) => {
            if (node.nodeType === Node.ELEMENT_NODE) {
              if (node.classList?.contains('ms-DetailsRow') || 
                  node.querySelector?.('.ms-DetailsRow') ||
                  node.classList?.contains('ms-List-cell')) {
                shouldRescan = true;
              }
            }
          });
        }
      });
      
      if (shouldRescan && !this.isProcessingAllPages) {
        clearTimeout(this.rescanTimeout);
        this.rescanTimeout = setTimeout(() => this.scanForInvoices(), 500);
      }
    });
    
    this.observer.observe(document.body, {
      childList: true,
      subtree: true
    });
  }
  
  scanForInvoices() {
    try {
      console.log('ETA Exporter: Starting invoice scan...');
      this.invoiceData = [];
      
      // Extract pagination info first
      this.extractPaginationInfo();
      
      // Find invoice rows using multiple methods
      const rows = this.getAllInvoiceRows();
      console.log(`ETA Exporter: Found ${rows.length} invoice rows on page ${this.currentPage}`);
      
      rows.forEach((row, index) => {
        const invoiceData = this.extractDataFromRow(row, index + 1);
        if (this.isValidInvoiceData(invoiceData)) {
          this.invoiceData.push(invoiceData);
        }
      });
      
      console.log(`ETA Exporter: Successfully extracted ${this.invoiceData.length} valid invoices from page ${this.currentPage}`);
      
    } catch (error) {
      console.error('ETA Exporter: Error scanning for invoices:', error);
    }
  }
  
  getAllInvoiceRows() {
    // Try multiple selectors to find invoice rows
    const selectors = [
      '.ms-DetailsRow[role="row"]',
      '.ms-List-cell[role="gridcell"]',
      '[data-list-index]',
      '.ms-DetailsRow',
      '[role="row"]',
      'tr[role="row"]',
      '.ms-List-cell',
      '[data-automation-key]',
      '.ms-DetailsRow-cell',
      'div[role="gridcell"]'
    ];
    
    let allRows = [];
    
    for (const selector of selectors) {
      try {
        const elements = document.querySelectorAll(selector);
        const validRows = Array.from(elements).filter(element => {
          // For cells, get the parent row
          const row = element.closest('[role="row"]') || element.parentElement || element;
          return this.isValidInvoiceRow(row) && !allRows.includes(row);
        }).map(element => element.closest('[role="row"]') || element.parentElement || element);
        
        allRows.push(...validRows);
        
        if (allRows.length > 0) {
          console.log(`ETA Exporter: Found ${validRows.length} rows using selector: ${selector}`);
        }
      } catch (error) {
        console.warn(`Error with selector ${selector}:`, error);
      }
    }
    
    // Remove duplicates
    allRows = [...new Set(allRows)];
    
    // Filter for visible and valid rows
    return allRows.filter(row => this.isRowVisible(row) && this.hasInvoiceData(row));
  }
  
  isValidInvoiceRow(row) {
    if (!row) return false;
    
    // Check if row has invoice-like data
    const text = row.textContent || '';
    
    // Look for patterns that indicate this is an invoice row
    const hasElectronicNumber = /[A-Z0-9]{15,}/.test(text);
    const hasAmount = /\d+[,٬]?\d*\.?\d*\s*(EGP|جنيه)/.test(text);
    const hasDate = /\d{1,2}\/\d{1,2}\/\d{4}/.test(text);
    const hasInvoiceElements = row.querySelector('.internalId-link, .griCellTitle, .griCellSubTitle');
    
    return hasElectronicNumber || hasAmount || hasDate || hasInvoiceElements;
  }
  
  isRowVisible(row) {
    if (!row) return false;
    
    const rect = row.getBoundingClientRect();
    const style = window.getComputedStyle(row);
    
    return (
      rect.width > 0 && 
      rect.height > 0 &&
      style.display !== 'none' &&
      style.visibility !== 'hidden' &&
      style.opacity !== '0'
    );
  }
  
  hasInvoiceData(row) {
    if (!row) return false;
    
    // Check for electronic number, internal number, or amount
    const electronicNumber = row.querySelector('.internalId-link a, [data-automation-key="uuid"] a, .griCellTitle');
    const internalNumber = row.querySelector('.griCellSubTitle, [data-automation-key="uuid"] .griCellSubTitle');
    const totalAmount = row.querySelector('[data-automation-key="total"], .griCellTitleGray');
    
    // Also check text content for patterns
    const text = row.textContent || '';
    const hasElectronicPattern = /[A-Z0-9]{15,}/.test(text);
    const hasAmountPattern = /\d+[,٬]?\d*\.?\d*\s*(EGP|جنيه)/.test(text);
    
    return !!(electronicNumber?.textContent?.trim() || 
              internalNumber?.textContent?.trim() || 
              totalAmount?.textContent?.trim() ||
              hasElectronicPattern ||
              hasAmountPattern);
  }
  
  extractPaginationInfo() {
    try {
      // Extract total count more aggressively
      this.totalCount = this.extractTotalCountImproved();
      
      // Extract current page
      this.currentPage = this.extractCurrentPageImproved();
      
      // Detect results per page
      this.resultsPerPage = this.detectResultsPerPageImproved();
      
      // Calculate total pages
      this.totalPages = Math.ceil(this.totalCount / this.resultsPerPage);
      
      // Ensure valid values
      this.currentPage = Math.max(this.currentPage, 1);
      this.totalPages = Math.max(this.totalPages, this.currentPage);
      this.totalCount = Math.max(this.totalCount, this.invoiceData.length);
      
      console.log(`ETA Exporter: Page ${this.currentPage} of ${this.totalPages}, Total: ${this.totalCount} invoices (${this.resultsPerPage} per page)`);
      
    } catch (error) {
      console.warn('ETA Exporter: Error extracting pagination info:', error);
      this.setDefaultPaginationValues();
    }
  }
  
  extractTotalCountImproved() {
    // Method 1: Look for "Results: XXX" pattern
    const allElements = document.querySelectorAll('*');
    
    for (const element of allElements) {
      const text = element.textContent || '';
      
      // Multiple patterns to match
      const patterns = [
        /Results:\s*(\d+)/i,
        /(\d+)\s*results/i,
        /total:\s*(\d+)/i,
        /النتائج:\s*(\d+)/i,
        /(\d+)\s*نتيجة/i,
        /من\s*(\d+)/i,
        /إجمالي:\s*(\d+)/i,
        /(\d+)\s*-\s*\d+\s*of\s*(\d+)/i,
        /(\d+)\s*-\s*\d+\s*من\s*(\d+)/i,
        /showing\s*\d+\s*-\s*\d+\s*of\s*(\d+)/i
      ];
      
      for (const pattern of patterns) {
        const match = text.match(pattern);
        if (match) {
          const count = parseInt(match[match.length - 1]); // Get last captured group
          if (count > 0 && count > this.invoiceData.length) {
            console.log(`ETA Exporter: Found total count ${count} using pattern: ${pattern}`);
            return count;
          }
        }
      }
    }
    
    // Method 2: Look in specific areas
    const paginationAreas = document.querySelectorAll('.ms-CommandBar, [class*="pagination"], [class*="pager"], .ms-Nav, [class*="footer"]');
    for (const area of paginationAreas) {
      const text = area.textContent || '';
      const match = text.match(/(\d+)\s*-\s*(\d+)\s*of\s*(\d+)|(\d+)\s*-\s*(\d+)\s*من\s*(\d+)/i);
      if (match) {
        const total = parseInt(match[3] || match[6]);
        if (total > 0) {
          console.log(`ETA Exporter: Found total count ${total} from pagination area`);
          return total;
        }
      }
    }
    
    // Method 3: Estimate from page numbers
    const maxPageNumber = this.findMaxPageNumber();
    if (maxPageNumber > 1) {
      const estimatedTotal = maxPageNumber * this.resultsPerPage;
      console.log(`ETA Exporter: Estimated total count ${estimatedTotal} from max page ${maxPageNumber}`);
      return estimatedTotal;
    }
    
    return Math.max(this.invoiceData.length, 50); // Default fallback
  }
  
  extractCurrentPageImproved() {
    // Method 1: Look for active page button
    const activeSelectors = [
      '.ms-Button--primary[aria-pressed="true"]',
      '[aria-current="page"]',
      '.active',
      '.selected',
      '.current',
      '[class*="active"][class*="page"]',
      '[class*="selected"][class*="page"]',
      '.ms-Button[aria-pressed="true"]'
    ];
    
    for (const selector of activeSelectors) {
      try {
        const activeButton = document.querySelector(selector);
        if (activeButton) {
          const pageText = activeButton.textContent?.trim();
          const pageNum = parseInt(pageText);
          if (!isNaN(pageNum) && pageNum > 0) {
            console.log(`ETA Exporter: Found current page ${pageNum} using active selector`);
            return pageNum;
          }
        }
      } catch (error) {
        // Continue to next selector
      }
    }
    
    // Method 2: Look for page input field
    const pageInputs = document.querySelectorAll('input[type="number"], input[aria-label*="page"], input[placeholder*="page"]');
    for (const input of pageInputs) {
      const value = parseInt(input.value);
      if (!isNaN(value) && value > 0) {
        console.log(`ETA Exporter: Found current page ${value} from input field`);
        return value;
      }
    }
    
    // Method 3: Look in URL
    const urlParams = new URLSearchParams(window.location.search);
    const pageFromUrl = urlParams.get('page') || urlParams.get('p') || urlParams.get('pageNumber');
    if (pageFromUrl) {
      const pageNum = parseInt(pageFromUrl);
      if (!isNaN(pageNum) && pageNum > 0) {
        console.log(`ETA Exporter: Found current page ${pageNum} from URL`);
        return pageNum;
      }
    }
    
    return 1; // Default to page 1
  }
  
  detectResultsPerPageImproved() {
    const currentPageRows = this.invoiceData.length;
    
    // Method 1: Look for page size selector
    const pageSizeSelectors = document.querySelectorAll('select, .ms-Dropdown-title');
    for (const selector of pageSizeSelectors) {
      const text = selector.textContent || '';
      const match = text.match(/(\d+)\s*(per page|items|عنصر)/i);
      if (match) {
        const size = parseInt(match[1]);
        if (size > 0) {
          console.log(`ETA Exporter: Found page size ${size} from selector`);
          return size;
        }
      }
    }
    
    // Method 2: Use current row count if not on last page
    if (this.currentPage < this.totalPages && currentPageRows > 0) {
      return currentPageRows;
    }
    
    // Method 3: Common page sizes
    const commonSizes = [10, 20, 25, 50, 100];
    for (const size of commonSizes) {
      if (currentPageRows <= size) {
        return size;
      }
    }
    
    return Math.max(currentPageRows, 50); // Default
  }
  
  findMaxPageNumber() {
    const pageButtons = document.querySelectorAll('button, a');
    let maxPage = 1;
    
    pageButtons.forEach(btn => {
      const buttonText = btn.textContent?.trim();
      const pageNum = parseInt(buttonText);
      
      if (!isNaN(pageNum) && pageNum > maxPage && pageNum < 1000) { // Reasonable upper limit
        maxPage = pageNum;
      }
    });
    
    return maxPage;
  }
  
  setDefaultPaginationValues() {
    this.currentPage = 1;
    this.totalPages = this.findMaxPageNumber() || 1;
    this.totalCount = Math.max(this.invoiceData.length, this.totalPages * 50);
    this.resultsPerPage = 50;
  }
  
  extractDataFromRow(row, index) {
    const invoice = {
      index: index,
      pageNumber: this.currentPage,
      
      // Main invoice data matching Excel format
      serialNumber: index,
      viewButton: 'عرض',
      documentType: 'فاتورة',
      documentVersion: '1.0',
      status: '',
      issueDate: '',
      submissionDate: '',
      invoiceCurrency: 'EGP',
      invoiceValue: '',
      vatAmount: '',
      taxDiscount: '0',
      totalInvoice: '',
      internalNumber: '',
      electronicNumber: '',
      sellerTaxNumber: '',
      sellerName: '',
      sellerAddress: '',
      buyerTaxNumber: '',
      buyerName: '',
      buyerAddress: '',
      purchaseOrderRef: '',
      purchaseOrderDesc: '',
      salesOrderRef: '',
      electronicSignature: 'موقع إلكترونياً',
      foodDrugGuide: '',
      externalLink: '',
      
      // Additional fields
      issueTime: '',
      totalAmount: '',
      currency: 'EGP',
      submissionId: '',
      details: []
    };
    
    try {
      // Extract data using multiple methods
      this.extractUsingDataAttributes(row, invoice);
      this.extractUsingCellPositions(row, invoice);
      this.extractUsingTextContent(row, invoice);
      
      // Generate external link if we have electronic number
      if (invoice.electronicNumber) {
        invoice.externalLink = this.generateExternalLink(invoice);
      }
      
    } catch (error) {
      console.warn(`ETA Exporter: Error extracting data from row ${index}:`, error);
    }
    
    return invoice;
  }
  
  extractUsingDataAttributes(row, invoice) {
    const cells = row.querySelectorAll('.ms-DetailsRow-cell, [data-automation-key]');
    
    cells.forEach(cell => {
      const key = cell.getAttribute('data-automation-key');
      
      switch (key) {
        case 'uuid':
          const electronicLink = cell.querySelector('.internalId-link a.griCellTitle, a');
          if (electronicLink) {
            invoice.electronicNumber = electronicLink.textContent?.trim() || '';
          }
          
          const internalNumberElement = cell.querySelector('.griCellSubTitle');
          if (internalNumberElement) {
            invoice.internalNumber = internalNumberElement.textContent?.trim() || '';
          }
          break;
          
        case 'dateTimeReceived':
          const dateElement = cell.querySelector('.griCellTitleGray, .griCellTitle');
          const timeElement = cell.querySelector('.griCellSubTitle');
          
          if (dateElement) {
            invoice.issueDate = dateElement.textContent?.trim() || '';
            invoice.submissionDate = invoice.issueDate;
          }
          if (timeElement) {
            invoice.issueTime = timeElement.textContent?.trim() || '';
          }
          break;
          
        case 'typeName':
          const typeElement = cell.querySelector('.griCellTitleGray, .griCellTitle');
          const versionElement = cell.querySelector('.griCellSubTitle');
          
          if (typeElement) {
            invoice.documentType = typeElement.textContent?.trim() || 'فاتورة';
          }
          if (versionElement) {
            invoice.documentVersion = versionElement.textContent?.trim() || '1.0';
          }
          break;
          
        case 'total':
          const totalElement = cell.querySelector('.griCellTitleGray, .griCellTitle');
          if (totalElement) {
            const totalText = totalElement.textContent?.trim() || '';
            invoice.totalAmount = totalText;
            invoice.totalInvoice = totalText;
            
            // Calculate VAT and invoice value
            const totalValue = this.parseAmount(totalText);
            if (totalValue > 0) {
              const vatRate = 0.14;
              const vatAmount = (totalValue * vatRate) / (1 + vatRate);
              const invoiceValue = totalValue - vatAmount;
              
              invoice.vatAmount = this.formatAmount(vatAmount);
              invoice.invoiceValue = this.formatAmount(invoiceValue);
            }
          }
          break;
          
        case 'issuerName':
          const sellerNameElement = cell.querySelector('.griCellTitleGray, .griCellTitle');
          const sellerTaxElement = cell.querySelector('.griCellSubTitle');
          
          if (sellerNameElement) {
            invoice.sellerName = sellerNameElement.textContent?.trim() || '';
          }
          if (sellerTaxElement) {
            invoice.sellerTaxNumber = sellerTaxElement.textContent?.trim() || '';
          }
          
          if (invoice.sellerName && !invoice.sellerAddress) {
            invoice.sellerAddress = 'غير محدد';
          }
          break;
          
        case 'receiverName':
          const buyerNameElement = cell.querySelector('.griCellTitleGray, .griCellTitle');
          const buyerTaxElement = cell.querySelector('.griCellSubTitle');
          
          if (buyerNameElement) {
            invoice.buyerName = buyerNameElement.textContent?.trim() || '';
          }
          if (buyerTaxElement) {
            invoice.buyerTaxNumber = buyerTaxElement.textContent?.trim() || '';
          }
          
          if (invoice.buyerName && !invoice.buyerAddress) {
            invoice.buyerAddress = 'غير محدد';
          }
          break;
          
        case 'submission':
          const submissionLink = cell.querySelector('a.submissionId-link, a');
          if (submissionLink) {
            invoice.submissionId = submissionLink.textContent?.trim() || '';
            invoice.purchaseOrderRef = invoice.submissionId;
          }
          break;
          
        case 'status':
          const validRejectedDiv = cell.querySelector('.horizontal.valid-rejected');
          if (validRejectedDiv) {
            const validStatus = validRejectedDiv.querySelector('.status-Valid');
            const rejectedStatus = validRejectedDiv.querySelector('.status-Rejected');
            if (validStatus && rejectedStatus) {
              invoice.status = `${validStatus.textContent?.trim()} → ${rejectedStatus.textContent?.trim()}`;
            }
          } else {
            const textStatus = cell.querySelector('.textStatus, .griCellTitle, .griCellTitleGray');
            if (textStatus) {
              invoice.status = textStatus.textContent?.trim() || '';
            }
          }
          break;
      }
    });
  }
  
  extractUsingCellPositions(row, invoice) {
    const cells = row.querySelectorAll('.ms-DetailsRow-cell, td, [role="gridcell"]');
    
    if (cells.length >= 6) {
      // Extract electronic number from first cell if not found
      if (!invoice.electronicNumber) {
        const firstCell = cells[0];
        const link = firstCell.querySelector('a');
        if (link) {
          invoice.electronicNumber = link.textContent?.trim() || '';
        }
      }
      
      // Extract total amount from cells
      if (!invoice.totalAmount) {
        for (let i = 2; i < Math.min(8, cells.length); i++) {
          const cellText = cells[i].textContent?.trim() || '';
          if (cellText.includes('EGP') || /^\d+[\d,]*\.?\d*$/.test(cellText.replace(/[,٬]/g, ''))) {
            invoice.totalAmount = cellText;
            invoice.totalInvoice = cellText;
            break;
          }
        }
      }
      
      // Extract date
      if (!invoice.issueDate) {
        for (let i = 1; i < Math.min(5, cells.length); i++) {
          const cellText = cells[i].textContent?.trim() || '';
          if (cellText.includes('/') && cellText.length >= 8) {
            invoice.issueDate = cellText;
            invoice.submissionDate = cellText;
            break;
          }
        }
      }
    }
  }
  
  extractUsingTextContent(row, invoice) {
    const allText = row.textContent || '';
    
    // Extract electronic number pattern
    if (!invoice.electronicNumber) {
      const electronicMatch = allText.match(/[A-Z0-9]{15,30}/);
      if (electronicMatch) {
        invoice.electronicNumber = electronicMatch[0];
      }
    }
    
    // Extract date pattern
    if (!invoice.issueDate) {
      const dateMatch = allText.match(/\d{1,2}\/\d{1,2}\/\d{4}/);
      if (dateMatch) {
        invoice.issueDate = dateMatch[0];
        invoice.submissionDate = dateMatch[0];
      }
    }
    
    // Extract amount pattern
    if (!invoice.totalAmount) {
      const amountMatch = allText.match(/\d+[,٬]?\d*\.?\d*\s*EGP/);
      if (amountMatch) {
        invoice.totalAmount = amountMatch[0];
        invoice.totalInvoice = amountMatch[0];
      }
    }
  }
  
  parseAmount(amountText) {
    if (!amountText) return 0;
    const cleanText = amountText.replace(/[,٬\sEGP]/g, '').replace(/[^\d.]/g, '');
    return parseFloat(cleanText) || 0;
  }
  
  formatAmount(amount) {
    if (!amount || amount === 0) return '0';
    return amount.toLocaleString('en-US', { 
      minimumFractionDigits: 2, 
      maximumFractionDigits: 2 
    });
  }
  
  generateExternalLink(invoice) {
    if (!invoice.electronicNumber) return '';
    
    let shareId = '';
    if (invoice.submissionId && invoice.submissionId.length > 10) {
      shareId = invoice.submissionId;
    } else {
      shareId = invoice.electronicNumber.replace(/[^A-Z0-9]/g, '').substring(0, 26);
    }
    
    return `https://invoicing.eta.gov.eg/documents/${invoice.electronicNumber}/share/${shareId}`;
  }
  
  isValidInvoiceData(invoice) {
    return !!(invoice.electronicNumber || invoice.internalNumber || invoice.totalAmount);
  }
  
  async getAllPagesData(options = {}) {
    try {
      this.isProcessingAllPages = true;
      this.allPagesData = [];
      
      console.log(`ETA Exporter: Starting to load ALL pages. Total invoices to load: ${this.totalCount}`);
      
      // Try the fastest method first - increase page size
      const fastResult = await this.tryIncreasePageSize();
      if (fastResult.success && fastResult.data.length >= this.totalCount * 0.9) {
        console.log(`ETA Exporter: Fast method successful! Loaded ${fastResult.data.length} invoices`);
        return fastResult;
      }
      
      // Fallback to optimized page navigation
      const result = await this.loadAllPagesOptimized(options);
      
      console.log(`ETA Exporter: Completed! Loaded ${result.data.length} invoices out of ${this.totalCount} total.`);
      
      return result;
      
    } catch (error) {
      console.error('ETA Exporter: Error getting all pages data:', error);
      return { 
        success: false, 
        data: this.allPagesData,
        error: error.message,
        totalProcessed: this.allPagesData.length
      };
    } finally {
      this.isProcessingAllPages = false;
    }
  }
  
  async tryIncreasePageSize() {
    try {
      console.log('ETA Exporter: Trying to increase page size for bulk loading...');
      
      // Look for page size controls
      const pageSizeControls = [
        'select[aria-label*="Items per page"]',
        'select[aria-label*="عدد العناصر"]',
        '.ms-Dropdown-title',
        '[data-automation-key="pageSize"]',
        'select:has(option[value="100"])',
        'select:has(option[value="200"])',
        '.ms-Dropdown'
      ];
      
      for (const selector of pageSizeControls) {
        const control = document.querySelector(selector);
        if (control) {
          console.log(`ETA Exporter: Found page size control: ${selector}`);
          
          if (control.tagName === 'SELECT') {
            // Handle select dropdown
            const options = control.querySelectorAll('option');
            let maxOption = null;
            let maxValue = 0;
            
            options.forEach(option => {
              const value = parseInt(option.value);
              if (!isNaN(value) && value > maxValue && value >= this.totalCount) {
                maxValue = value;
                maxOption = option;
              }
            });
            
            // If no option covers all, get the largest available
            if (!maxOption) {
              options.forEach(option => {
                const value = parseInt(option.value);
                if (!isNaN(value) && value > maxValue) {
                  maxValue = value;
                  maxOption = option;
                }
              });
            }
            
            if (maxOption && maxValue > this.resultsPerPage) {
              console.log(`ETA Exporter: Setting page size to ${maxValue}`);
              control.value = maxOption.value;
              control.dispatchEvent(new Event('change', { bubbles: true }));
              
              // Wait for page to reload
              await this.delay(3000);
              await this.waitForPageLoadComplete();
              
              // Scan for all invoices
              this.scanForInvoices();
              
              if (this.invoiceData.length >= this.totalCount * 0.8) {
                const processedData = this.invoiceData.map((invoice, index) => ({
                  ...invoice,
                  pageNumber: 1,
                  serialNumber: index + 1,
                  globalIndex: index + 1
                }));
                
                return { success: true, data: processedData };
              }
            }
          } else if (control.classList.contains('ms-Dropdown')) {
            // Handle Microsoft Fabric dropdown
            control.click();
            await this.delay(500);
            
            const dropdownOptions = document.querySelectorAll('.ms-Dropdown-item, [role="option"]');
            let bestOption = null;
            let bestValue = 0;
            
            dropdownOptions.forEach(option => {
              const text = option.textContent?.trim() || '';
              const value = parseInt(text);
              if (!isNaN(value) && value > bestValue) {
                bestValue = value;
                bestOption = option;
              }
            });
            
            if (bestOption && bestValue > this.resultsPerPage) {
              console.log(`ETA Exporter: Clicking option for ${bestValue} items`);
              bestOption.click();
              
              await this.delay(3000);
              await this.waitForPageLoadComplete();
              
              this.scanForInvoices();
              
              if (this.invoiceData.length >= this.totalCount * 0.8) {
                const processedData = this.invoiceData.map((invoice, index) => ({
                  ...invoice,
                  pageNumber: 1,
                  serialNumber: index + 1,
                  globalIndex: index + 1
                }));
                
                return { success: true, data: processedData };
              }
            }
          }
        }
      }
      
      return { success: false, error: 'Could not increase page size' };
      
    } catch (error) {
      console.error('Error increasing page size:', error);
      return { success: false, error: error.message };
    }
  }
  
  async loadAllPagesOptimized(options = {}) {
    try {
      console.log('ETA Exporter: Using optimized page-by-page loading...');
      
      const allData = [];
      let processedInvoices = 0;
      
      // Start from page 1 for consistency
      await this.navigateToPage(1);
      await this.waitForPageLoadComplete();
      
      // Process all pages
      for (let pageNum = 1; pageNum <= this.totalPages; pageNum++) {
        try {
          if (this.progressCallback) {
            this.progressCallback({
              currentPage: pageNum,
              totalPages: this.totalPages,
              message: `جاري معالجة الصفحة ${pageNum} من ${this.totalPages}... (${processedInvoices}/${this.totalCount} فاتورة)`,
              percentage: this.totalPages > 0 ? (pageNum / this.totalPages) * 100 : 0
            });
          }
          
          // Navigate to page if not already there
          if (this.currentPage !== pageNum) {
            const navigated = await this.navigateToPage(pageNum);
            if (!navigated) {
              console.warn(`Could not navigate to page ${pageNum}`);
              continue;
            }
            
            await this.delay(1000);
            await this.waitForPageLoadComplete();
          }
          
          // Scan current page
          this.scanForInvoices();
          
          if (this.invoiceData.length > 0) {
            const pageData = this.invoiceData.map((invoice, index) => ({
              ...invoice,
              pageNumber: pageNum,
              serialNumber: ((pageNum - 1) * this.resultsPerPage) + index + 1,
              globalIndex: processedInvoices + index + 1
            }));
            
            allData.push(...pageData);
            processedInvoices += this.invoiceData.length;
            
            console.log(`ETA Exporter: Page ${pageNum} processed, collected ${this.invoiceData.length} invoices. Total: ${processedInvoices}/${this.totalCount}`);
          }
          
          // Check if we have enough data
          if (processedInvoices >= this.totalCount) {
            console.log('ETA Exporter: Collected all expected invoices, stopping early');
            break;
          }
          
        } catch (error) {
          console.error(`Error processing page ${pageNum}:`, error);
          continue;
        }
      }
      
      return { 
        success: true, 
        data: allData,
        totalProcessed: processedInvoices,
        expectedTotal: this.totalCount
      };
      
    } catch (error) {
      console.error('Optimized loading failed:', error);
      return { success: false, error: error.message, data: [] };
    }
  }
  
  async navigateToPage(pageNumber) {
    try {
      console.log(`ETA Exporter: Navigating to page ${pageNumber}`);
      
      // Method 1: Click page number button directly
      const pageButtons = document.querySelectorAll('button, a');
      for (const button of pageButtons) {
        const buttonText = button.textContent?.trim();
        if (parseInt(buttonText) === pageNumber && !button.disabled) {
          console.log(`ETA Exporter: Clicking page ${pageNumber} button`);
          button.click();
          this.currentPage = pageNumber;
          return true;
        }
      }
      
      // Method 2: Use page input field
      const pageInputs = document.querySelectorAll('input[type="number"], input[aria-label*="page"], input[placeholder*="page"]');
      for (const input of pageInputs) {
        console.log(`ETA Exporter: Using page input to go to page ${pageNumber}`);
        input.value = pageNumber.toString();
        input.dispatchEvent(new Event('input', { bubbles: true }));
        input.dispatchEvent(new Event('change', { bubbles: true }));
        input.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));
        this.currentPage = pageNumber;
        return true;
      }
      
      // Method 3: Navigate step by step if close
      const currentPage = this.extractCurrentPageImproved();
      const diff = pageNumber - currentPage;
      
      if (Math.abs(diff) <= 5) {
        console.log(`ETA Exporter: Navigating step by step from page ${currentPage} to ${pageNumber}`);
        for (let i = 0; i < Math.abs(diff); i++) {
          const success = diff > 0 ? await this.navigateToNextPage() : await this.navigateToPreviousPage();
          if (!success) return false;
          await this.delay(800);
        }
        return true;
      }
      
      console.warn(`ETA Exporter: Could not navigate to page ${pageNumber}`);
      return false;
      
    } catch (error) {
      console.error(`Error navigating to page ${pageNumber}:`, error);
      return false;
    }
  }
  
  async navigateToNextPage() {
    const nextButton = this.findNextButton();
    if (nextButton && !nextButton.disabled) {
      nextButton.click();
      this.currentPage++;
      return true;
    }
    return false;
  }
  
  async navigateToPreviousPage() {
    const prevButton = this.findPreviousButton();
    if (prevButton && !prevButton.disabled) {
      prevButton.click();
      this.currentPage--;
      return true;
    }
    return false;
  }
  
  findNextButton() {
    const nextSelectors = [
      'button[aria-label*="Next"]',
      'button[aria-label*="next"]',
      'button[title*="Next"]',
      'button[title*="next"]',
      'button[aria-label*="التالي"]',
      'button:has([data-icon-name="ChevronRight"])',
      'button:has([class*="chevron-right"])',
      '.ms-Button:has([data-icon-name="ChevronRight"])'
    ];
    
    for (const selector of nextSelectors) {
      try {
        const button = document.querySelector(selector);
        if (button && !button.disabled) {
          return button;
        }
      } catch (e) {
        // Continue
      }
    }
    
    // Fallback: look for buttons with ">" or "»"
    const allButtons = document.querySelectorAll('button');
    for (const button of allButtons) {
      if (button.disabled) continue;
      const text = button.textContent?.trim() || '';
      if (text === '>' || text === '»' || text.includes('next') || text.includes('التالي')) {
        return button;
      }
    }
    
    return null;
  }
  
  findPreviousButton() {
    const prevSelectors = [
      'button[aria-label*="Previous"]',
      'button[aria-label*="previous"]',
      'button[title*="Previous"]',
      'button[title*="previous"]',
      'button[aria-label*="السابق"]',
      'button:has([data-icon-name="ChevronLeft"])',
      'button:has([class*="chevron-left"])',
      '.ms-Button:has([data-icon-name="ChevronLeft"])'
    ];
    
    for (const selector of prevSelectors) {
      try {
        const button = document.querySelector(selector);
        if (button && !button.disabled) {
          return button;
        }
      } catch (e) {
        // Continue
      }
    }
    
    // Fallback: look for buttons with "<" or "«"
    const allButtons = document.querySelectorAll('button');
    for (const button of allButtons) {
      if (button.disabled) continue;
      const text = button.textContent?.trim() || '';
      if (text === '<' || text === '«' || text.includes('previous') || text.includes('السابق')) {
        return button;
      }
    }
    
    return null;
  }
  
  async waitForPageLoadComplete() {
    console.log('ETA Exporter: Waiting for page load to complete...');
    
    // Wait for loading indicators to disappear
    await this.waitForCondition(() => {
      const loadingIndicators = document.querySelectorAll(
        '.LoadingIndicator, .ms-Spinner, [class*="loading"], [class*="spinner"], .ms-Shimmer'
      );
      const isLoading = Array.from(loadingIndicators).some(el => 
        el.offsetParent !== null && 
        window.getComputedStyle(el).display !== 'none'
      );
      return !isLoading;
    }, 15000);
    
    // Wait for invoice rows to appear
    await this.waitForCondition(() => {
      const rows = this.getAllInvoiceRows();
      return rows.length > 0;
    }, 15000);
    
    // Additional stability wait
    await this.delay(1000);
    
    console.log('ETA Exporter: Page load completed');
  }
  
  async waitForCondition(condition, timeout = 10000) {
    const startTime = Date.now();
    
    while (Date.now() - startTime < timeout) {
      try {
        if (condition()) {
          return true;
        }
      } catch (error) {
        // Ignore errors in condition check
      }
      await this.delay(300);
    }
    
    console.warn(`ETA Exporter: Condition timeout after ${timeout}ms`);
    return false;
  }
  
  delay(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
  }
  
  setProgressCallback(callback) {
    this.progressCallback = callback;
  }
  
  async getInvoiceDetails(invoiceId) {
    try {
      const details = await this.extractInvoiceDetailsFromPage(invoiceId);
      return {
        success: true,
        data: details
      };
    } catch (error) {
      console.error('Error getting invoice details:', error);
      return { 
        success: false, 
        data: [],
        error: error.message 
      };
    }
  }
  
  async extractInvoiceDetailsFromPage(invoiceId) {
    const details = [];
    
    try {
      const detailsTable = document.querySelector('.ms-DetailsList, [data-automationid="DetailsList"], table');
      
      if (detailsTable) {
        const rows = detailsTable.querySelectorAll('.ms-DetailsRow[role="row"], tr');
        
        rows.forEach((row, index) => {
          const cells = row.querySelectorAll('.ms-DetailsRow-cell, td');
          
          if (cells.length >= 6) {
            const item = {
              itemCode: this.extractCellText(cells[0]) || `ITEM-${index + 1}`,
              description: this.extractCellText(cells[1]) || 'صنف',
              unitCode: this.extractCellText(cells[2]) || 'EA',
              unitName: this.extractCellText(cells[3]) || 'قطعة',
              quantity: this.extractCellText(cells[4]) || '1',
              unitPrice: this.extractCellText(cells[5]) || '0',
              totalValue: this.extractCellText(cells[6]) || '0',
              taxAmount: this.extractCellText(cells[7]) || '0',
              vatAmount: this.extractCellText(cells[8]) || '0'
            };
            
            if (item.description && 
                item.description !== 'اسم الصنف' && 
                item.description !== 'Description' &&
                item.description.trim() !== '') {
              details.push(item);
            }
          }
        });
      }
      
      if (details.length === 0) {
        const invoice = this.invoiceData.find(inv => inv.electronicNumber === invoiceId);
        if (invoice) {
          details.push({
            itemCode: invoice.electronicNumber || 'INVOICE',
            description: 'إجمالي الفاتورة',
            unitCode: 'EA',
            unitName: 'فاتورة',
            quantity: '1',
            unitPrice: invoice.totalAmount || '0',
            totalValue: invoice.invoiceValue || invoice.totalAmount || '0',
            taxAmount: '0',
            vatAmount: invoice.vatAmount || '0'
          });
        }
      }
      
    } catch (error) {
      console.error('Error extracting invoice details:', error);
    }
    
    return details;
  }
  
  extractCellText(cell) {
    if (!cell) return '';
    
    const textElement = cell.querySelector('.griCellTitle, .griCellTitleGray, .ms-DetailsRow-cellContent') || cell;
    return textElement.textContent?.trim() || '';
  }
  
  getInvoiceData() {
    return {
      invoices: this.invoiceData,
      totalCount: this.totalCount,
      currentPage: this.currentPage,
      totalPages: this.totalPages
    };
  }
  
  cleanup() {
    if (this.observer) {
      this.observer.disconnect();
    }
    if (this.rescanTimeout) {
      clearTimeout(this.rescanTimeout);
    }
  }
}

// Initialize content script
const etaContentScript = new ETAContentScript();

// Listen for messages from popup
chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {
  console.log('ETA Exporter: Received message:', request.action);
  
  switch (request.action) {
    case 'ping':
      sendResponse({ success: true, message: 'Content script is ready' });
      break;
      
    case 'getInvoiceData':
      const data = etaContentScript.getInvoiceData();
      console.log('ETA Exporter: Returning invoice data:', data);
      sendResponse({
        success: true,
        data: data
      });
      break;
      
    case 'getInvoiceDetails':
      etaContentScript.getInvoiceDetails(request.invoiceId)
        .then(result => sendResponse(result))
        .catch(error => sendResponse({ success: false, error: error.message }));
      return true;
      
    case 'getAllPagesData':
      if (request.options && request.options.progressCallback) {
        etaContentScript.setProgressCallback((progress) => {
          chrome.runtime.sendMessage({
            action: 'progressUpdate',
            progress: progress
          }).catch(() => {
            // Ignore errors if popup is closed
          });
        });
      }
      
      etaContentScript.getAllPagesData(request.options)
        .then(result => {
          console.log('ETA Exporter: All pages data result:', result);
          sendResponse(result);
        })
        .catch(error => {
          console.error('ETA Exporter: Error in getAllPagesData:', error);
          sendResponse({ success: false, error: error.message });
        });
      return true;
      
    case 'rescanPage':
      etaContentScript.scanForInvoices();
      sendResponse({
        success: true,
        data: etaContentScript.getInvoiceData()
      });
      break;
      
    default:
      sendResponse({ success: false, error: 'Unknown action' });
  }
  
  return true;
});

// Cleanup on page unload
window.addEventListener('beforeunload', () => {
  etaContentScript.cleanup();
});

console.log('ETA Exporter: Content script loaded successfully');