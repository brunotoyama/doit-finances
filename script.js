// CDN Health Check - verifica bibliotecas após carregamento
(function() {
  var CDN_TIMEOUT = 10000; // 10 segundos
  var libs = [
    { name: 'Chart.js', global: 'Chart' },
    { name: 'SheetJS',  global: 'XLSX' },
    { name: 'jsPDF',    global: 'jspdf' }
  ];

  setTimeout(function() {
    var missing = libs.filter(function(lib) {
      return typeof window[lib.global] === 'undefined';
    });

    if (missing.length > 0) {
      var names = missing.map(function(l) { return l.name; }).join(', ');
      var banner = document.createElement('div');
      banner.className = 'cdn-error-banner';
      banner.setAttribute('role', 'alert');
      banner.innerHTML = '⚠️ Falha ao carregar: <strong>' + names + 
        '</strong>. Verifique sua conexão com a internet e recarregue a página.';
      document.body.insertBefore(banner, document.body.firstChild);
    }
  }, CDN_TIMEOUT);
})();

/**
 * Financial Dashboard - DOit ERP Excel Analysis Panel
 *
 * Single-file application logic organized into logical sections.
 * Uses Chart.js (CDN), SheetJS/xlsx (CDN), and jsPDF (CDN).
 * Event-driven architecture with inline EventBus for inter-module communication.
 *
 * @file script.js
 */

// ============================================================
// SECTION: Configuration & Constants
// ============================================================

/**
 * @typedef {Object} FinancialRecord
 * @property {string} id - Generated UUID for internal tracking
 * @property {string} projeto - Project name
 * @property {string} cliente - Client name
 * @property {string} categoria - Category
 * @property {number} valor - Monetary value in BRL
 * @property {string} tipo - 'receita' or 'despesa'
 * @property {Date} data - Transaction date
 * @property {string} status - 'pago'|'recebido'|'pendente'|'vencido'
 * @property {string} centroCusto - Cost center
 * @property {string} responsavel - Responsible person
 * @property {string} departamento - Department
 * @property {string} conta - Account
 */

/**
 * @typedef {Object} FilterCriteria
 * @property {Date} [dataInicio] - Start date
 * @property {Date} [dataFim] - End date
 * @property {string} [projeto] - Project filter
 * @property {string} [cliente] - Client filter
 * @property {string} [categoria] - Category filter
 * @property {string} [centroCusto] - Cost center filter
 * @property {string} [status] - Status filter
 * @property {string} [responsavel] - Responsible filter
 */

/**
 * @typedef {Object} KPIResult
 * @property {number} totalRecebido - Sum of received payments
 * @property {number} totalPago - Sum of made payments
 * @property {number} saldo - totalRecebido - totalPago
 * @property {number} recebimentosAberto - Sum of pending receivables
 * @property {number} pagamentosAberto - Sum of pending payables
 * @property {number} fluxoMes - Net cash flow for current month
 * @property {number} ticketMedio - Average transaction value
 * @property {number} quantidadeLancamentos - Total entry count
 * @property {Object} variacoes - Month-over-month % variations per KPI
 */

/**
 * @typedef {Object} InsightResult
 * @property {Array} expenseIncreases - Expenses up >10% MoM
 * @property {Array} revenueDrops - Revenue down >5% MoM
 * @property {Array} topClients - Top 5 clients by revenue
 * @property {Array} topCategories - Top 5 expense categories
 * @property {Array} monthlyVariations - All month comparisons
 * @property {Array} alerts - Overdue payment alerts
 * @property {string} trendDirection - '↑'|'↓'|'→' 3-month moving direction
 * @property {Array} projectRanking - Top 10 projects by net value
 * @property {Array} clientRanking - Top 10 clients by net value
 * @property {Array} categoryRanking - Top 10 categories by net value
 * @property {Object|null} forecast - 3-month projection or null
 */

/**
 * Expected columns with their variations for matching.
 * Each entry maps a canonical field name to possible header variations.
 */
const COLUMN_DEFINITIONS = {
  idProjeto:     { required: false, aliases: ['id projeto'] },
  projeto:       { required: false, aliases: ['projeto', 'project', 'proj'] },
  idCliente:     { required: false, aliases: ['id de / para', 'id de/para'] },
  cliente:       { required: false, aliases: ['de / para', 'de/para', 'cliente', 'client', 'cli', 'de / para (apelido)'] },
  tipoPagamento: { required: false, aliases: ['tipo'] },
  tipoClassif:   { required: false, aliases: ['1a categoria', '1ª categoria'] },
  nivel1:        { required: false, aliases: ['2a categoria', '2ª categoria'] },
  nivel2:        { required: false, aliases: ['3a categoria', '3ª categoria'] },
  categoria:     { required: false, aliases: ['categoria', 'category', 'cat'] },
  valor:         { required: false, aliases: ['valor', 'value', 'val', 'montante', 'previsto'] },
  receita:       { required: false, aliases: ['+'] },
  despesa:       { required: false, aliases: ['-'] },
  data:          { required: true,  aliases: ['emissao', 'emissão', 'data', 'date', 'dt', 'data lancamento', 'vcto'] },
  status:        { required: false, aliases: ['conciliado', 'status', 'situacao', 'sit'] },
  centroCusto:   { required: false, aliases: ['departamento', 'depto', 'dept', 'dep'] },
  conta:         { required: false, aliases: ['conta', 'account', 'cta'] },
  descricao:     { required: false, aliases: ['descricao', 'descrição', 'descricão', 'historico', 'histórico'] }
};

/**
 * Chart configurations for all dashboard charts.
 * Includes dedicated charts for receitas and despesas by classification level.
 * @type {Array<{id: string, type: string, title: string, grouped?: boolean, horizontal?: boolean}>}
 */
const CHART_CONFIGS = [
  { id: 'fluxo-caixa',          type: 'bar',      title: 'Fluxo de Caixa Mensal' },
  { id: 'evolucao-mensal',       type: 'line',     title: 'Evolução Financeira Mensal' },
  { id: 'receitas-chart',        type: 'doughnut', title: 'Receitas por Categoria' },
  { id: 'despesas-chart',        type: 'doughnut', title: 'Despesas por Categoria' },
  { id: 'recebimentos-proj',     type: 'bar',      title: 'Recebimentos por Projeto', horizontal: true },
  { id: 'top-clientes',          type: 'bar',      title: 'Top Clientes', horizontal: true },
  { id: 'tipo-pagamento',        type: 'bar',      title: 'Distribuição por Tipo de Pagamento', horizontal: true }
];

/**
 * Responsive breakpoint configuration.
 * Defines grid layouts for different viewport widths.
 */
const BREAKPOINTS = {
  notebook:  { min: 1024, max: 1366, kpiCols: 2, chartCols: 1 },
  desktop:   { min: 1367, max: 1920, kpiCols: 4, chartCols: 2 },
  ultrawide: { min: 1921, max: 3840, kpiCols: 4, chartCols: 2 }
};

/**
 * Performance tuning constants.
 */
const PERFORMANCE = {
  CHUNK_SIZE: 500,
  CHUNK_DELAY: 0,
  MAX_CHART_POINTS: 1000,
  MAX_ROWS: 100000,
  LAZY_RENDER_OFFSET: 200,
  ANIMATION_DURATION: 750,
  FILTER_DEBOUNCE: 100,
  SEARCH_DEBOUNCE: 200,
  PAGE_SIZE: 20
};

/**
 * Loading stage definitions for progress overlay.
 * Each stage has a label and weight for progress calculation.
 */
const LOADING_STAGES = [
  { id: 'reading',    label: 'Lendo arquivo...',         weight: 20 },
  { id: 'mapping',    label: 'Mapeando colunas...',      weight: 10 },
  { id: 'processing', label: 'Processando dados...',     weight: 40 },
  { id: 'rendering',  label: 'Renderizando dashboard...', weight: 30 }
];

// ============================================================
// SECTION: EventBus (Inline Implementation)
// ============================================================

/**
 * Lightweight publish/subscribe event system for inter-module communication.
 * Mirrors the API of src/utils/event-bus.js for consistency.
 *
 * Events used by the Financial Dashboard:
 * - 'data:loaded'       - Data imported and ready
 * - 'filters:changed'   - Filter criteria updated
 * - 'filters:cleared'   - All filters reset
 * - 'export:requested'  - PDF/image export triggered
 * - 'section:navigated' - Sidebar navigation occurred
 * - 'chart:clicked'     - Chart element clicked (drill-down)
 * - 'loading:start'     - Long operation started
 * - 'loading:end'       - Long operation completed
 * - 'error:occurred'    - Error requiring user feedback
 */
const EventBus = (() => {
  const listeners = {};

  return {
    /**
     * Register an event handler.
     * @param {string} event - Event name
     * @param {Function} handler - Callback function
     */
    on(event, handler) {
      if (typeof event !== 'string' || !event) {
        throw new Error('Event name must be a non-empty string');
      }
      if (typeof handler !== 'function') {
        throw new Error('Handler must be a function');
      }
      if (!listeners[event]) {
        listeners[event] = [];
      }
      listeners[event].push(handler);
    },

    /**
     * Remove a previously registered event handler.
     * @param {string} event - Event name
     * @param {Function} handler - Handler reference to remove
     */
    off(event, handler) {
      if (typeof event !== 'string' || !event) {
        throw new Error('Event name must be a non-empty string');
      }
      if (typeof handler !== 'function') {
        throw new Error('Handler must be a function');
      }
      if (!listeners[event]) return;
      listeners[event] = listeners[event].filter(h => h !== handler);
      if (listeners[event].length === 0) {
        delete listeners[event];
      }
    },

    /**
     * Emit an event, invoking all registered handlers.
     * @param {string} event - Event name
     * @param {*} data - Data to pass to handlers
     */
    emit(event, data) {
      if (typeof event !== 'string' || !event) {
        throw new Error('Event name must be a non-empty string');
      }
      if (!listeners[event]) return;
      const handlers = [...listeners[event]];
      for (const handler of handlers) {
        handler(data);
      }
    },

    /**
     * Remove all listeners for a specific event, or all listeners.
     * @param {string} [event] - Optional event name
     */
    clear(event) {
      if (event) {
        delete listeners[event];
      } else {
        for (const key of Object.keys(listeners)) {
          delete listeners[key];
        }
      }
    }
  };
})();

// ============================================================
// SECTION: Data Abstraction Layer
// ============================================================

/** @type {FinancialRecord[]} Module-scoped storage for all loaded records */
let _records = [];

/** @type {FilterCriteria} Module-scoped storage for active filter criteria */
let _activeFilters = {};

/** @type {Object} DataLayer - Data access interface for future API readiness */
const DataLayer = {
  /**
   * Get all loaded records (excluding marked records by default).
   * @param {boolean} [includeExcluded=false] - If true, include excluded records
   * @returns {FinancialRecord[]} All loaded records
   */
  getData(includeExcluded) {
    if (includeExcluded) return _records;
    return _records.filter(r => !r._excluded);
  },

  /**
   * Get records filtered by the given criteria using AND logic.
   * If criteria is empty or null/undefined, returns all records.
   * @param {FilterCriteria} criteria - Filter criteria to apply
   * @returns {FinancialRecord[]} Filtered records
   */
  getFilteredData(criteria) {
    const baseRecords = _records.filter(r => !r._excluded);
    if (!criteria || Object.keys(criteria).length === 0) {
      return baseRecords;
    }

    return baseRecords.filter(record => {
      // Date range filtering
      if (criteria.dataInicio) {
        const startDate = criteria.dataInicio instanceof Date ? criteria.dataInicio : new Date(criteria.dataInicio);
        if (record.data < startDate) return false;
      }
      if (criteria.dataFim) {
        const endDate = criteria.dataFim instanceof Date ? criteria.dataFim : new Date(criteria.dataFim);
        if (record.data > endDate) return false;
      }

      // String equality filters (AND logic)
      if (criteria.projeto && record.projeto !== criteria.projeto) return false;
      if (criteria.cliente && record.cliente !== criteria.cliente) return false;
      if (criteria.categoria && record.categoria !== criteria.categoria) return false;
      if (criteria.tipoClassif && record.tipoClassif !== criteria.tipoClassif) return false;
      if (criteria.tipoPagamento && record.tipoPagamento !== criteria.tipoPagamento) return false;
      if (criteria.nivel1 && record.nivel1 !== criteria.nivel1) return false;
      if (criteria.nivel2 && record.nivel2 !== criteria.nivel2) return false;
      if (criteria.centroCusto && record.centroCusto !== criteria.centroCusto) return false;
      if (criteria.conta && record.conta !== criteria.conta) return false;
      if (criteria.status && record.status !== criteria.status) return false;

      return true;
    });
  },

  /**
   * Get sorted unique non-null values for a given column.
   * @param {string} columnName - The column name to extract values from
   * @returns {string[]} Sorted unique non-null values
   */
  getColumnValues(columnName) {
    const values = new Set();
    for (const record of _records) {
      if (record._excluded) continue;
      const value = record[columnName];
      if (value != null && value !== '') {
        values.add(value);
      }
    }
    return [...values].sort((a, b) => {
      if (typeof a === 'string' && typeof b === 'string') {
        return a.localeCompare(b, 'pt-BR');
      }
      return a < b ? -1 : a > b ? 1 : 0;
    });
  },

  /**
   * Set the data source and emit 'data:loaded' event.
   * @param {FinancialRecord[]} records - Array of financial records to store
   */
  setData(records) {
    _records = records;
    const columns = Object.keys(COLUMN_DEFINITIONS);
    const metadata = {
      totalRecords: records.length,
      importedAt: new Date()
    };
    EventBus.emit('data:loaded', { records, columns, metadata });
  },

  /**
   * Get the current active filter criteria.
   * @returns {FilterCriteria} Current active filters
   */
  getActiveFilters() {
    return _activeFilters;
  },

  /**
   * Apply filter criteria and emit 'filters:changed' event.
   * @param {FilterCriteria} criteria - Filter criteria to apply
   */
  setFilters(criteria) {
    _activeFilters = criteria || {};
    const filteredData = this.getFilteredData(_activeFilters);
    EventBus.emit('filters:changed', { criteria: _activeFilters, filteredData });
  },

  /**
   * Reset all filters and emit 'filters:cleared' event.
   */
  clearFilters() {
    _activeFilters = {};
    EventBus.emit('filters:cleared', {});
  }
};

// ============================================================
// SECTION: Excel Parser & Column Mapper
// ============================================================

/** @type {Object} ColumnMapper - Column identification and mapping */
const ColumnMapper = {
  /**
   * Normalize a header string for comparison.
   * Converts to lowercase, applies NFD decomposition, strips diacritical marks,
   * and trims whitespace.
   * @param {string} header - Raw header text
   * @returns {string} Normalized header
   */
  normalizeHeader(header) {
    if (typeof header !== 'string') return '';
    return header
      .trim()
      .toLowerCase()
      .normalize('NFD')
      .replace(/[\u0300-\u036f]/g, '');
  },

  /**
   * Map Excel headers to expected financial fields.
   * Uses case-insensitive, accent-insensitive matching against COLUMN_DEFINITIONS aliases.
   * @param {string[]} headers - Raw column headers from Excel
   * @returns {{mapped: Object<string,number>, unmapped: string[], missing: string[]}}
   */
  mapColumns(headers) {
    const mapped = {};
    const unmapped = [];
    const matchedIndices = new Set();

    // Normalize all aliases from COLUMN_DEFINITIONS for comparison
    const normalizedDefs = {};
    for (const [fieldName, def] of Object.entries(COLUMN_DEFINITIONS)) {
      normalizedDefs[fieldName] = def.aliases.map(alias => this.normalizeHeader(alias));
    }

    // For each header, try to find a matching field
    for (let i = 0; i < headers.length; i++) {
      const normalizedHeader = this.normalizeHeader(headers[i]);
      if (!normalizedHeader) {
        unmapped.push(headers[i]);
        continue;
      }

      let found = false;
      for (const [fieldName, normalizedAliases] of Object.entries(normalizedDefs)) {
        // Skip if this field is already mapped
        if (fieldName in mapped) continue;

        if (normalizedAliases.includes(normalizedHeader)) {
          mapped[fieldName] = i;
          matchedIndices.add(i);
          found = true;
          break;
        }
      }

      if (!found) {
        unmapped.push(headers[i]);
      }
    }

    // Determine missing fields (all fields not in mapped)
    const missing = Object.keys(COLUMN_DEFINITIONS).filter(field => !(field in mapped));

    // Special handling for DOit ERP: find the second "Realizado" column (numeric values)
    // The first "Realizado" is a date (column B), the second is the monetary value (column J)
    const realizadoIndices = [];
    for (let i = 0; i < headers.length; i++) {
      if (this.normalizeHeader(headers[i]) === 'realizado') {
        realizadoIndices.push(i);
      }
    }
    if (realizadoIndices.length >= 2) {
      // The second "Realizado" is the numeric value column
      mapped._realizadoNumerico = realizadoIndices[1];
    } else if (realizadoIndices.length === 1 && !('data' in mapped)) {
      // If only one "Realizado" and no date mapped, it might be the value
      mapped._realizadoNumerico = realizadoIndices[0];
    }

    return { mapped, unmapped, missing };
  },

  /**
   * Check if required columns (Valor, Tipo, Data) are present in the mapping.
   * @param {Object<string,number>} mapping - Column mapping result (the mapped object)
   * @returns {boolean}
   */
  hasRequiredColumns(mapping) {
    // For DOit ERP: we need at least a date column AND either valor/receita/despesa columns
    const hasDate = 'data' in mapping;
    const hasValue = 'valor' in mapping || 'receita' in mapping || 'despesa' in mapping;
    return hasDate && hasValue;
  },

  /**
   * Identify which sheets in a workbook contain financial data.
   * Scans all sheets and returns those whose first row headers match
   * at least one column definition alias.
   * @param {Object} workbook - SheetJS workbook object
   * @returns {string[]} Sheet names containing financial headers
   */
  identifyFinancialSheets(workbook) {
    const financialSheets = [];

    for (const sheetName of workbook.SheetNames) {
      const sheet = workbook.Sheets[sheetName];
      if (!sheet) continue;

      // Read headers from the first row using SheetJS utilities
      const range = XLSX.utils.decode_range(sheet['!ref'] || 'A1');
      const headers = [];

      for (let col = range.s.c; col <= range.e.c; col++) {
        const cellAddress = XLSX.utils.encode_cell({ r: range.s.r, c: col });
        const cell = sheet[cellAddress];
        if (cell && cell.v != null) {
          headers.push(String(cell.v));
        }
      }

      if (headers.length === 0) continue;

      // Check if any header matches a column definition
      const { mapped } = this.mapColumns(headers);
      if (Object.keys(mapped).length > 0) {
        financialSheets.push(sheetName);
      }
    }

    return financialSheets;
  }
};

/** @type {Object} ExcelParser - File parsing with chunked processing */
const ExcelParser = {
  /**
   * Generate a simple unique ID for each record.
   * Uses a combination of timestamp, random number, and counter.
   * @returns {string} Unique identifier string
   */
  _counter: 0,
  generateId() {
    this._counter++;
    return `rec_${Date.now().toString(36)}_${Math.random().toString(36).substr(2, 6)}_${this._counter}`;
  },

  /**
   * Parse a numeric value, handling Brazilian format (comma as decimal separator).
   * Handles strings like "1.234,56" or "1234.56" or plain numbers.
   * @param {*} raw - Raw value from Excel cell
   * @returns {number} Parsed numeric value, or 0 if unparseable
   */
  parseValue(raw) {
    if (raw == null || raw === '') return 0;
    if (typeof raw === 'number') return raw;

    const str = String(raw).trim();
    // Remove currency symbols and spaces
    const cleaned = str.replace(/[R$\s]/g, '');

    // Detect Brazilian format: has comma as decimal separator
    // Pattern: digits with optional dots (thousands) and comma (decimal)
    if (/^\-?[\d.]+,\d{1,2}$/.test(cleaned)) {
      // Brazilian format: 1.234,56 -> 1234.56
      const normalized = cleaned.replace(/\./g, '').replace(',', '.');
      const result = parseFloat(normalized);
      return isNaN(result) ? 0 : result;
    }

    // Try standard format
    const result = parseFloat(cleaned.replace(/,/g, ''));
    return isNaN(result) ? 0 : result;
  },

  /**
   * Parse a date value from various formats.
   * Handles: DD/MM/YYYY, YYYY-MM-DD, Excel serial numbers, Date objects.
   * @param {*} raw - Raw date value from Excel cell
   * @returns {Date} Parsed Date object, or current date if unparseable
   */
  parseDate(raw) {
    if (raw == null || raw === '') return new Date();
    if (raw instanceof Date) return raw;

    // Excel serial number (number of days since 1900-01-01)
    if (typeof raw === 'number') {
      // Excel date serial number conversion
      const excelEpoch = new Date(1899, 11, 30);
      const date = new Date(excelEpoch.getTime() + raw * 86400000);
      if (!isNaN(date.getTime())) return date;
      return new Date();
    }

    const str = String(raw).trim();

    // DD/MM/YYYY format
    const brMatch = str.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{4})$/);
    if (brMatch) {
      const [, day, month, year] = brMatch;
      const date = new Date(parseInt(year), parseInt(month) - 1, parseInt(day));
      if (!isNaN(date.getTime())) return date;
    }

    // YYYY-MM-DD format
    const isoMatch = str.match(/^(\d{4})[\/\-.](\d{1,2})[\/\-.](\d{1,2})/);
    if (isoMatch) {
      const [, year, month, day] = isoMatch;
      const date = new Date(parseInt(year), parseInt(month) - 1, parseInt(day));
      if (!isNaN(date.getTime())) return date;
    }

    // Try native Date parsing as fallback
    const fallback = new Date(str);
    if (!isNaN(fallback.getTime())) return fallback;

    return new Date();
  },

  /**
   * Normalize tipo value to 'receita' or 'despesa'.
   * Handles various representations including abbreviations and English terms.
   * @param {*} raw - Raw tipo value
   * @returns {string} 'receita' or 'despesa'
   */
  normalizeTipo(raw) {
    if (raw == null || raw === '') return 'despesa';
    const str = String(raw).trim().toLowerCase()
      .normalize('NFD').replace(/[\u0300-\u036f]/g, '');

    // Match receita variations (including DOit ERP types)
    if (/^(receitas?|revenue|income|rec|r|entrada|credit|credito|cr|lucro operacional|receitas nao operacionais)$/i.test(str)) {
      return 'receita';
    }

    // Match despesa variations (including DOit ERP types)
    if (/^(despesas?|expense|desp|d|saida|saidas?|debit|debito|db|gasto|custos?|pagamento|mao de obra|saidas nao operacionais)$/i.test(str)) {
      return 'despesa';
    }

    // DOit ERP specific: check if contains key words
    if (str.includes('receita') || str.includes('lucro')) return 'receita';
    if (str.includes('despesa') || str.includes('custo') || str.includes('saida')) return 'despesa';

    // Default to despesa if unrecognized
    return 'despesa';
  },

  /**
   * Normalize status value to one of: 'pago', 'recebido', 'pendente', 'vencido'.
   * @param {*} raw - Raw status value
   * @returns {string} 'Pago'|'Pendente'
   */
  normalizeStatus(raw) {
    if (raw == null || raw === '') return 'Pendente';
    const str = String(raw).trim().toLowerCase()
      .normalize('NFD').replace(/[\u0300-\u036f]/g, '');

    // DOit ERP uses "Conciliado" column with TRUE/FALSE
    if (str === 'true' || str === 'sim' || str === 'yes' || str === '1') {
      return 'Pago';
    }
    if (str === 'false' || str === 'nao' || str === 'no' || str === '0') {
      return 'Pendente';
    }

    // Match pago variations
    if (/^(pago|paid|pg|liquidado|quitado|recebido|received|rec|confirmado)$/i.test(str)) {
      return 'Pago';
    }

    // Match pendente variations
    if (/^(pendente|pending|pend|aberto|em aberto|aguardando|a receber|a pagar|vencido|overdue|atrasado)$/i.test(str)) {
      return 'Pendente';
    }

    // Default to Pendente if unrecognized
    return 'Pendente';
  },

  /**
   * Process rows in chunks to avoid blocking the main thread.
   * Each chunk processes at most PERFORMANCE.CHUNK_SIZE rows, yielding control
   * between chunks via setTimeout(0).
   * @param {Array<Array>} rows - Raw row data (array of arrays)
   * @param {Object} mapping - Column mapping from ColumnMapper.mapColumns (the mapped object)
   * @param {function} [onProgress] - Progress callback (stage, percent)
   * @returns {Promise<FinancialRecord[]>} Normalized financial records
   */
  async processChunked(rows, mapping, onProgress) {
    const records = [];
    const totalRows = Math.min(rows.length, PERFORMANCE.MAX_ROWS);
    const chunkSize = PERFORMANCE.CHUNK_SIZE;

    if (rows.length > PERFORMANCE.MAX_ROWS) {
      console.warn(`File contains ${rows.length} rows. Processing only the first ${PERFORMANCE.MAX_ROWS} rows.`);
    }

    for (let i = 0; i < totalRows; i += chunkSize) {
      const end = Math.min(i + chunkSize, totalRows);

      for (let j = i; j < end; j++) {
        const row = rows[j];
        if (!row || row.length === 0) continue;

        // Determine valor and tipo
        // Strategy: Use "Previsto"/"Valor" column sign as primary tipo indicator
        // Positive = receita, Negative = despesa
        // Fallback to "+" and "-" columns if Previsto is not available
        let valor = 0;
        let tipo = 'despesa';
        
        // Primary: Use "Previsto"/"Valor" column - its sign determines receita/despesa
        if (mapping.valor != null) {
          const rawVal = this.parseValue(row[mapping.valor]);
          if (rawVal !== 0) {
            valor = Math.abs(rawVal);
            tipo = rawVal > 0 ? 'receita' : 'despesa';
          }
        }

        // Fallback: use "+" column for receitas (if Previsto didn't have a value)
        if (valor === 0 && mapping.receita != null) {
          const recVal = this.parseValue(row[mapping.receita]);
          if (recVal > 0) {
            valor = recVal;
            tipo = 'receita';
          }
        }
        // Fallback: use "-" column for despesas
        if (valor === 0 && mapping.despesa != null) {
          const despVal = this.parseValue(row[mapping.despesa]);
          if (despVal > 0) {
            valor = despVal;
            tipo = 'despesa';
          }
        }

        // DEBUG: Log first 5 records to console for troubleshooting
        if (j < 5) {
          console.log(`[DEBUG] Row ${j}: receita_col=${mapping.receita}, despesa_col=${mapping.despesa}, raw_receita=${row[mapping.receita]}, raw_despesa=${row[mapping.despesa]}, valor=${valor}, tipo=${tipo}, projeto=${row[mapping.projeto]}, cliente=${row[mapping.cliente]}`);
        }
        // Last resort: scan for the second "Realizado" column (numeric one)
        // In DOit ERP, column J (index 9) is the numeric "Realizado" with signed values
        if (valor === 0 && mapping._realizadoNumerico != null) {
          const rawVal = this.parseValue(row[mapping._realizadoNumerico]);
          if (rawVal !== 0) {
            valor = Math.abs(rawVal);
            tipo = rawVal > 0 ? 'receita' : 'despesa';
          }
        }

        const record = {
          id: this.generateId(),
          idProjeto: mapping.idProjeto != null ? String(row[mapping.idProjeto] || '') : '',
          projeto: mapping.projeto != null ? String(row[mapping.projeto] || '') : '',
          idCliente: mapping.idCliente != null ? String(row[mapping.idCliente] || '') : '',
          cliente: mapping.cliente != null ? String(row[mapping.cliente] || '') : '',
          categoria: mapping.categoria != null ? String(row[mapping.categoria] || '') : '',
          tipoClassif: mapping.tipoClassif != null ? String(row[mapping.tipoClassif] || '') : '',
          tipoPagamento: mapping.tipoPagamento != null ? String(row[mapping.tipoPagamento] || '') : '',
          nivel1: mapping.nivel1 != null ? String(row[mapping.nivel1] || '') : '',
          nivel2: mapping.nivel2 != null ? String(row[mapping.nivel2] || '') : '',
          valor: valor,
          tipo: tipo,
          data: mapping.data != null ? this.parseDate(row[mapping.data]) : new Date(),
          status: mapping.status != null ? this.normalizeStatus(row[mapping.status]) : 'Pendente',
          centroCusto: mapping.centroCusto != null ? String(row[mapping.centroCusto] || '') : '',
          conta: mapping.conta != null ? String(row[mapping.conta] || '') : '',
          descricao: mapping.descricao != null ? String(row[mapping.descricao] || '') : ''
        };

        // tipo is already correctly determined from +/- columns above
        // tipoClassif (1ª Categoria) is for classification display only, not for tipo determination

        // Use nivel1 as categoria if categoria is empty
        if (!record.categoria && record.nivel1) {
          record.categoria = record.nivel1;
        }

        // Mark excluded records but still include them
        record._excluded = false;
        if (valor === 0) {
          record._excluded = true;
          record._excludeReason = 'Valor zero';
        }
        // Saldo Inicial is NOT excluded - it's included in all calculations
        // but marked for identification

        records.push(record);
      }

      // Report progress during processing stage (60-80% range)
      if (onProgress) {
        const progress = 60 + Math.round((end / totalRows) * 20);
        onProgress('processing', progress);
      }

      // Yield to the main thread between chunks
      if (end < totalRows) {
        await new Promise(r => setTimeout(r, PERFORMANCE.CHUNK_DELAY));
      }
    }

    return records;
  },

  /**
   * Parse an Excel file using chunked processing.
   * Reads the file, identifies financial sheets, maps columns, and processes rows.
   * @param {File} file - The Excel file to parse
   * @param {function} [onProgress] - Progress callback (stage, percent)
   * @returns {Promise<{records: FinancialRecord[], metadata: Object}>}
   */
  async parseFile(file, onProgress) {
    // Emit loading start
    EventBus.emit('loading:start', { stage: 'reading' });

    try {
      // Step 1: Read file as ArrayBuffer
      const data = await new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = (e) => resolve(e.target.result);
        reader.onerror = (e) => reject(new Error('Erro ao ler o arquivo: ' + (e.target.error?.message || 'erro desconhecido')));
        reader.readAsArrayBuffer(file);
      });

      if (onProgress) onProgress('reading', 20);

      // Step 2: Parse with SheetJS
      const workbook = XLSX.read(data, { type: 'array' });

      // Step 3: Identify financial sheets
      const financialSheets = ColumnMapper.identifyFinancialSheets(workbook);

      if (financialSheets.length === 0) {
        throw new Error('Nenhuma planilha com dados financeiros reconhecíveis foi encontrada. Verifique se o arquivo contém colunas como Valor, Tipo e Data.');
      }

      if (onProgress) onProgress('mapping', 40);

      // Step 4: Process each financial sheet
      let allRecords = [];
      const sheetMetadata = [];

      for (let s = 0; s < financialSheets.length; s++) {
        const sheetName = financialSheets[s];
        const sheet = workbook.Sheets[sheetName];

        // Convert sheet to array of arrays (header: 1 gives raw arrays)
        const rawData = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });

        if (rawData.length < 2) continue; // Need at least header + 1 data row

        // First row is headers
        const headers = rawData[0].map(h => String(h || ''));
        let { mapped, unmapped, missing } = ColumnMapper.mapColumns(headers);

        // DEBUG: Log column mapping result
        console.log('[DEBUG] Headers:', headers);
        console.log('[DEBUG] Mapped columns:', JSON.stringify(mapped));
        console.log('[DEBUG] Unmapped:', unmapped);

        // Check required columns - if missing, show manual mapping UI
        if (!ColumnMapper.hasRequiredColumns(mapped)) {
          // Pause loading overlay while user maps columns
          EventBus.emit('loading:end', {});

          // Determine which required fields are missing
          const requiredFields = Object.entries(COLUMN_DEFINITIONS)
            .filter(([, def]) => def.required)
            .map(([fieldName]) => fieldName);
          const missingRequired = requiredFields.filter(f => !(f in mapped));

          // All file columns (not just unmapped) should be available for selection
          const allFileColumns = headers.filter(h => h && h.trim() !== '');

          try {
            // Show column mapping UI and wait for user response
            const manualMapping = await UIController.showColumnMapping(allFileColumns, missingRequired);

            // Merge manual mapping with auto-mapped columns
            // manualMapping is {fieldName: columnName}, convert to {fieldName: columnIndex}
            for (const [fieldName, columnName] of Object.entries(manualMapping)) {
              const colIndex = headers.indexOf(columnName);
              if (colIndex !== -1) {
                mapped[fieldName] = colIndex;
              }
            }

            // Re-check required columns after manual mapping
            if (!ColumnMapper.hasRequiredColumns(mapped)) {
              console.warn(`Sheet "${sheetName}" still missing required columns after manual mapping, skipping.`);
              continue;
            }

            // Resume loading overlay
            EventBus.emit('loading:start', { stage: 'processing' });
          } catch (cancelError) {
            // User cancelled - throw to abort the entire parsing
            throw new Error('Mapeamento cancelado pelo usuário.');
          }
        }

        // Data rows (skip header)
        const dataRows = rawData.slice(1);

        if (onProgress) onProgress('processing', 60);

        // Process rows in chunks
        const records = await this.processChunked(dataRows, mapped, onProgress);

        allRecords = allRecords.concat(records);
        sheetMetadata.push({
          name: sheetName,
          rowCount: dataRows.length,
          recordCount: records.length,
          columns: Object.keys(mapped)
        });
      }

      if (onProgress) onProgress('rendering', 80);

      // Step 5: Store results via DataLayer
      DataLayer.setData(allRecords);

      // DEBUG: Summary of parsed data
      const receitas = allRecords.filter(r => r.tipo === 'receita');
      const despesas = allRecords.filter(r => r.tipo === 'despesa');
      const totalReceita = receitas.reduce((s, r) => s + r.valor, 0);
      const totalDespesa = despesas.reduce((s, r) => s + r.valor, 0);
      console.log(`[DEBUG] Total records: ${allRecords.length}, Receitas: ${receitas.length} (R$ ${totalReceita.toFixed(2)}), Despesas: ${despesas.length} (R$ ${totalDespesa.toFixed(2)})`);
      if (receitas.length > 0) {
        console.log('[DEBUG] Sample receita:', JSON.stringify(receitas[0]));
      }
      // Top projetos
      const projTotals = {};
      receitas.forEach(r => { if (r.projeto) projTotals[r.projeto] = (projTotals[r.projeto] || 0) + r.valor; });
      const topProj = Object.entries(projTotals).sort((a,b) => b[1]-a[1]).slice(0,5);
      console.log('[DEBUG] Top 5 projetos:', topProj.map(([n,v]) => `${n}: ${v.toFixed(2)}`).join(', '));

      const metadata = {
        fileName: file.name,
        fileSize: file.size,
        sheetsProcessed: sheetMetadata.length,
        totalRecords: allRecords.length,
        sheets: sheetMetadata,
        importedAt: new Date()
      };

      // Emit loading end
      EventBus.emit('loading:end', {});

      return { records: allRecords, metadata };

    } catch (error) {
      // Emit loading end even on error
      EventBus.emit('loading:end', {});
      throw error;
    }
  }
};

// ============================================================
// SECTION: Filter Engine
// ============================================================

/** @type {Object} FilterEngine - AND-logic filter application */
const FilterEngine = {
  /** @type {number|null} Internal timer reference for debouncing */
  _debounceTimer: null,

  /**
   * Apply AND-logic filters to a dataset.
   * Each active criterion must be satisfied for a record to be included.
   * If criteria is empty/null/undefined, returns all data unchanged.
   *
   * @param {FinancialRecord[]} data - Full dataset
   * @param {FilterCriteria} criteria - Filter criteria
   * @returns {FinancialRecord[]} Filtered records
   */
  applyFilters(data, criteria) {
    if (!data || !Array.isArray(data)) return [];
    if (!criteria || Object.keys(criteria).length === 0) return data;

    return data.filter(record => {
      // Date range filtering: exclude records before dataInicio
      if (criteria.dataInicio) {
        const startDate = criteria.dataInicio instanceof Date
          ? criteria.dataInicio
          : new Date(criteria.dataInicio);
        if (record.data < startDate) return false;
      }

      // Date range filtering: exclude records after dataFim
      if (criteria.dataFim) {
        const endDate = criteria.dataFim instanceof Date
          ? criteria.dataFim
          : new Date(criteria.dataFim);
        if (record.data > endDate) return false;
      }

      // String equality filters (AND logic - all must match)
      if (criteria.projeto && record.projeto !== criteria.projeto) return false;
      if (criteria.cliente && record.cliente !== criteria.cliente) return false;
      if (criteria.categoria && record.categoria !== criteria.categoria) return false;
      if (criteria.tipoClassif && record.tipoClassif !== criteria.tipoClassif) return false;
      if (criteria.tipoPagamento && record.tipoPagamento !== criteria.tipoPagamento) return false;
      if (criteria.nivel1 && record.nivel1 !== criteria.nivel1) return false;
      if (criteria.nivel2 && record.nivel2 !== criteria.nivel2) return false;
      if (criteria.centroCusto && record.centroCusto !== criteria.centroCusto) return false;
      if (criteria.conta && record.conta !== criteria.conta) return false;
      if (criteria.status && record.status !== criteria.status) return false;

      return true;
    });
  },

  /**
   * Extract unique non-null values for a given column from the dataset.
   * Returns a sorted array with no duplicates.
   *
   * @param {FinancialRecord[]} data - Dataset
   * @param {string} column - Column name
   * @returns {string[]} Sorted unique values
   */
  getUniqueValues(data, column) {
    if (!data || !Array.isArray(data) || !column) return [];

    const values = new Set();
    for (const record of data) {
      const value = record[column];
      if (value != null && value !== '') {
        values.add(value);
      }
    }

    return [...values].sort((a, b) => {
      if (typeof a === 'string' && typeof b === 'string') {
        return a.localeCompare(b, 'pt-BR');
      }
      return a < b ? -1 : a > b ? 1 : 0;
    });
  },

  /**
   * Debounced version of filter application.
   * Clears any previous pending timer and sets a new one with
   * PERFORMANCE.FILTER_DEBOUNCE (100ms) delay. When the timer fires,
   * calls DataLayer.setFilters(criteria).
   *
   * @param {FilterCriteria} criteria - Filter criteria to apply after debounce
   */
  applyFiltersDebounced(criteria) {
    if (this._debounceTimer) {
      clearTimeout(this._debounceTimer);
    }
    this._debounceTimer = setTimeout(() => {
      DataLayer.setFilters(criteria);
      this._debounceTimer = null;
    }, PERFORMANCE.FILTER_DEBOUNCE);
  }
};

// ============================================================
// SECTION: KPI Calculator
// ============================================================

/** @type {Object} KPICalculator - Financial KPI computation */
const KPICalculator = {
  /**
   * Calculate all KPI values from a dataset.
   * @param {FinancialRecord[]} data - Financial records
   * @returns {KPIResult} All computed KPIs
   */
  calculateAll(data) {
    if (!data || !Array.isArray(data) || data.length === 0) {
      return {
        totalRecebido: 0,
        totalPago: 0,
        saldo: 0,
        recebimentosAberto: 0,
        pagamentosAberto: 0,
        fluxoMes: 0,
        ticketMedio: 0,
        quantidadeLancamentos: 0,
        variacoes: {}
      };
    }

    // totalRecebido: sum of valor where tipo='receita' AND status='Pago'
    const totalRecebido = data
      .filter(r => r.tipo === 'receita' && r.status === 'Pago')
      .reduce((sum, r) => sum + r.valor, 0);

    // totalPago: sum of valor where tipo='despesa' AND status='Pago'
    const totalPago = data
      .filter(r => r.tipo === 'despesa' && r.status === 'Pago')
      .reduce((sum, r) => sum + r.valor, 0);

    // saldo: totalRecebido - totalPago
    const saldo = totalRecebido - totalPago;

    // recebimentosAberto: sum of valor where tipo='receita' AND status='Pendente'
    const recebimentosAberto = data
      .filter(r => r.tipo === 'receita' && r.status === 'Pendente')
      .reduce((sum, r) => sum + r.valor, 0);

    // pagamentosAberto: sum of valor where tipo='despesa' AND status='Pendente'
    const pagamentosAberto = data
      .filter(r => r.tipo === 'despesa' && r.status === 'Pendente')
      .reduce((sum, r) => sum + r.valor, 0);

    // fluxoMes: net cash flow for the most recent month in the dataset
    const fluxoMes = this._calculateFluxoMes(data);

    // ticketMedio: total valor / number of records
    const totalValor = data.reduce((sum, r) => sum + r.valor, 0);
    const ticketMedio = data.length > 0 ? totalValor / data.length : 0;

    // quantidadeLancamentos: total count of ALL records (including excluded)
    const quantidadeLancamentos = _records ? _records.length : data.length;

    // variacoes: percentage variations for each KPI
    const variacoes = {
      totalRecebido: this.calculateVariation(data, 'totalRecebido'),
      totalPago: this.calculateVariation(data, 'totalPago'),
      saldo: this.calculateVariation(data, 'saldo'),
      recebimentosAberto: this.calculateVariation(data, 'recebimentosAberto'),
      pagamentosAberto: this.calculateVariation(data, 'pagamentosAberto'),
      fluxoMes: this.calculateVariation(data, 'fluxoMes'),
      ticketMedio: this.calculateVariation(data, 'ticketMedio'),
      quantidadeLancamentos: this.calculateVariation(data, 'quantidadeLancamentos')
    };

    return {
      totalRecebido,
      totalPago,
      saldo,
      recebimentosAberto,
      pagamentosAberto,
      fluxoMes,
      ticketMedio,
      quantidadeLancamentos,
      variacoes
    };
  },

  /**
   * Calculate net cash flow for the most recent month in the dataset.
   * @param {FinancialRecord[]} data - Financial records
   * @returns {number} Net cash flow (receitas - despesas) for the most recent month
   * @private
   */
  _calculateFluxoMes(data) {
    if (!data || data.length === 0) return 0;

    // Find the most recent month in the dataset
    let latestDate = data[0].data;
    for (const record of data) {
      if (record.data > latestDate) {
        latestDate = record.data;
      }
    }

    const latestYear = latestDate.getFullYear();
    const latestMonth = latestDate.getMonth();

    // Filter records for the most recent month
    const monthRecords = data.filter(r => {
      return r.data.getFullYear() === latestYear && r.data.getMonth() === latestMonth;
    });

    // Net cash flow: receitas - despesas for that month
    const receitas = monthRecords
      .filter(r => r.tipo === 'receita')
      .reduce((sum, r) => sum + r.valor, 0);
    const despesas = monthRecords
      .filter(r => r.tipo === 'despesa')
      .reduce((sum, r) => sum + r.valor, 0);

    return receitas - despesas;
  },

  /**
   * Group data by month and compute a specific KPI for each month.
   * @param {FinancialRecord[]} data - Financial records
   * @param {string} kpiName - KPI identifier
   * @returns {Map<string, number>} Map of 'YYYY-MM' -> KPI value
   * @private
   */
  _getMonthlyKPI(data, kpiName) {
    // Group records by month (YYYY-MM)
    const monthGroups = new Map();
    for (const record of data) {
      const key = `${record.data.getFullYear()}-${String(record.data.getMonth() + 1).padStart(2, '0')}`;
      if (!monthGroups.has(key)) {
        monthGroups.set(key, []);
      }
      monthGroups.get(key).push(record);
    }

    // Compute KPI for each month
    const monthlyValues = new Map();
    for (const [monthKey, records] of monthGroups) {
      let value = 0;
      switch (kpiName) {
        case 'totalRecebido':
          value = records
            .filter(r => r.tipo === 'receita' && r.status === 'Pago')
            .reduce((sum, r) => sum + r.valor, 0);
          break;
        case 'totalPago':
          value = records
            .filter(r => r.tipo === 'despesa' && r.status === 'Pago')
            .reduce((sum, r) => sum + r.valor, 0);
          break;
        case 'saldo': {
          const rec = records
            .filter(r => r.tipo === 'receita' && r.status === 'Pago')
            .reduce((sum, r) => sum + r.valor, 0);
          const pag = records
            .filter(r => r.tipo === 'despesa' && r.status === 'Pago')
            .reduce((sum, r) => sum + r.valor, 0);
          value = rec - pag;
          break;
        }
        case 'recebimentosAberto':
          value = records
            .filter(r => r.tipo === 'receita' && r.status === 'Pendente')
            .reduce((sum, r) => sum + r.valor, 0);
          break;
        case 'pagamentosAberto':
          value = records
            .filter(r => r.tipo === 'despesa' && r.status === 'Pendente')
            .reduce((sum, r) => sum + r.valor, 0);
          break;
        case 'fluxoMes': {
          const receitas = records
            .filter(r => r.tipo === 'receita')
            .reduce((sum, r) => sum + r.valor, 0);
          const despesas = records
            .filter(r => r.tipo === 'despesa')
            .reduce((sum, r) => sum + r.valor, 0);
          value = receitas - despesas;
          break;
        }
        case 'ticketMedio': {
          const total = records.reduce((sum, r) => sum + r.valor, 0);
          value = records.length > 0 ? total / records.length : 0;
          break;
        }
        case 'quantidadeLancamentos':
          value = records.length;
          break;
        default:
          value = 0;
      }
      monthlyValues.set(monthKey, value);
    }

    return monthlyValues;
  },

  /**
   * Calculate month-over-month percentage variation for a KPI.
   * ((currentMonth - previousMonth) / |previousMonth|) * 100
   * @param {FinancialRecord[]} data - Financial records
   * @param {string} kpiName - KPI identifier
   * @returns {number|null} Percentage change or null if insufficient data
   */
  calculateVariation(data, kpiName) {
    if (!data || !Array.isArray(data) || data.length === 0) return null;

    const monthlyValues = this._getMonthlyKPI(data, kpiName);

    // Need at least 2 months of data
    if (monthlyValues.size < 2) return null;

    // Sort months chronologically
    const sortedMonths = [...monthlyValues.keys()].sort();
    const currentMonthKey = sortedMonths[sortedMonths.length - 1];
    const previousMonthKey = sortedMonths[sortedMonths.length - 2];

    const currentValue = monthlyValues.get(currentMonthKey);
    const previousValue = monthlyValues.get(previousMonthKey);

    // Cannot compute variation if previous month is 0
    if (previousValue === 0) return null;

    return ((currentValue - previousValue) / Math.abs(previousValue)) * 100;
  },

  /**
   * Format a number as BRL currency (R$ #.###,00).
   * Uses period as thousands separator and comma as decimal separator.
   * @param {number} value - Numeric value
   * @returns {string} Formatted BRL string
   */
  formatBRL(value) {
    if (value == null || isNaN(value)) return 'R$ 0,00';

    const isNegative = value < 0;
    const absValue = Math.abs(value);

    // Format with 2 decimal places
    const fixed = absValue.toFixed(2);
    const [intPart, decPart] = fixed.split('.');

    // Add period as thousands separator
    const withThousands = intPart.replace(/\B(?=(\d{3})+(?!\d))/g, '.');

    // Compose final string
    const formatted = `${withThousands},${decPart}`;
    return isNegative ? `R$ -${formatted}` : `R$ ${formatted}`;
  },

  /**
   * Format a number with period as thousands separator (integer format).
   * @param {number} value - Integer value
   * @returns {string} Formatted integer string
   */
  formatInteger(value) {
    if (value == null || isNaN(value)) return '0';

    const intValue = Math.round(value);
    const isNegative = intValue < 0;
    const absStr = String(Math.abs(intValue));

    // Add period as thousands separator
    const withThousands = absStr.replace(/\B(?=(\d{3})+(?!\d))/g, '.');

    return isNegative ? `-${withThousands}` : withThousands;
  },

  /**
   * Determine display color based on value sign.
   * @param {number} value - Numeric value
   * @returns {'positive'|'negative'|'neutral'} Color class
   */
  getValueColor(value) {
    if (value > 0) return 'positive';
    if (value < 0) return 'negative';
    return 'neutral';
  }
};

// ============================================================
// SECTION: Chart Engine
// ============================================================

/** @type {Object} ChartEngine - Chart.js chart management */
const ChartEngine = {
  /** @type {Object<string, Chart>} Chart.js instances keyed by chart ID */
  _instances: {},

  /** @type {string[]} Professional financial color palette for series */
  _colors: ['#1d4ed8', '#0d7a3f', '#b45309', '#7c3aed', '#0891b2', '#be185d', '#4338ca', '#15803d', '#c2410c', '#6d28d9'],

  /** @type {string[]} 12 visually distinct colors for month-based bar charts */
  _monthPalette: ['#1d4ed8', '#0d9488', '#d97706', '#7c3aed', '#0891b2', '#be185d', '#15803d', '#c2410c', '#4338ca', '#b45309', '#0e7490', '#9333ea'],

  /** @type {string} Color for positive values */
  _positiveColor: '#0d7a3f',

  /** @type {string} Color for negative values */
  _negativeColor: '#c41e3a',

  /**
   * Initialize all 10 charts with Chart.js instances.
   * Creates a Chart.js instance for each entry in CHART_CONFIGS.
   * @param {FinancialRecord[]} data - Financial records
   */
  initCharts(data) {
    // Destroy any existing chart instances
    for (const id of Object.keys(this._instances)) {
      if (this._instances[id]) {
        this._instances[id].destroy();
      }
    }
    this._instances = {};

    for (const config of CHART_CONFIGS) {
      const canvas = document.getElementById(`chart-${config.id}`);
      if (!canvas) continue;

      const ctx = canvas.getContext('2d');
      const chartData = this._prepareChartData(config, data);
      const options = this._getChartOptions(config);

      // Determine Chart.js chart type
      let chartType = config.type;
      if (config.type === 'bar' && config.horizontal) {
        chartType = 'bar'; // horizontal is handled via indexAxis in options
      }

      // Add onClick handler for chart interactivity (drill-down filtering)
      const chartId = config.id;
      options.onClick = (event, elements, chart) => {
        // If clicked on a bar/slice element, use it directly
        if (elements && elements.length > 0) {
          const element = elements[0];
          const filter = this.getFilterFromClick(chartId, element);
          if (filter && Object.keys(filter).length > 0) {
            EventBus.emit('chart:clicked', { chartId, filter });
          }
          return;
        }

        // For horizontal bar charts, check if click was on a Y-axis label
        if (config.horizontal && chart.scales && chart.scales.y) {
          const yScale = chart.scales.y;
          const nativeEvent = event.native || event;
          const rect = chart.canvas.getBoundingClientRect();
          const x = nativeEvent.clientX - rect.left;
          const y = nativeEvent.clientY - rect.top;
          
          // Check if click is in the label area (left of chart area)
          if (x < chart.chartArea.left + 5) {
            // Find which label was clicked based on Y position
            for (let i = 0; i < chart.data.labels.length; i++) {
              const labelY = yScale.getPixelForValue(i);
              const tickHeight = yScale.height / chart.data.labels.length;
              if (Math.abs(y - labelY) < tickHeight / 2) {
                const fakeElement = { index: i };
                const filter = this.getFilterFromClick(chartId, fakeElement);
                if (filter && Object.keys(filter).length > 0) {
                  EventBus.emit('chart:clicked', { chartId, filter });
                }
                return;
              }
            }
          }
        }
      };

      this._instances[config.id] = new Chart(ctx, {
        type: chartType,
        data: chartData,
        options: options
      });

      // Add native click listener for Y-axis labels on horizontal bar charts
      if (config.horizontal) {
        const chartRef = this._instances[config.id];
        const chartIdRef = config.id;
        const self = this;
        canvas.addEventListener('click', (e) => {
          const rect = canvas.getBoundingClientRect();
          const x = e.clientX - rect.left;
          const y = e.clientY - rect.top;
          
          // Only act if click is in the Y-axis label area
          if (!chartRef.chartArea || x >= chartRef.chartArea.left) return;
          
          const yScale = chartRef.scales.y;
          if (!yScale) return;

          // Find which label index matches the click Y position
          for (let i = 0; i < chartRef.data.labels.length; i++) {
            const labelY = yScale.getPixelForValue(i);
            const tickSpacing = yScale.height / chartRef.data.labels.length;
            if (Math.abs(y - labelY) <= tickSpacing / 2) {
              const fakeElement = { index: i };
              const filter = self.getFilterFromClick(chartIdRef, fakeElement);
              if (filter && Object.keys(filter).length > 0) {
                EventBus.emit('chart:clicked', { chartId: chartIdRef, filter });
              }
              break;
            }
          }
        });
      }
    }
  },

  /**
   * Update all existing chart instances with new data.
   * Shows empty state message when chart has no data after filtering.
   * @param {FinancialRecord[]} data - Filtered financial records
   */
  updateCharts(data) {
    for (const config of CHART_CONFIGS) {
      const chart = this._instances[config.id];
      if (!chart) continue;

      const chartData = this._prepareChartData(config, data);
      chart.data = chartData;
      chart.update('active');

      // Empty state handling: check if chart has no data
      const container = chart.canvas.closest('.chart-container');
      if (container) {
        const canvas = chart.canvas;
        const emptyMsg = container.querySelector('.chart-empty');
        const hasData = chartData.labels && chartData.labels.length > 0 &&
          chartData.datasets && chartData.datasets.length > 0 &&
          chartData.datasets.some(ds => ds.data && ds.data.some(v => v !== 0));

        if (!hasData) {
          canvas.style.display = 'none';
          if (emptyMsg) emptyMsg.hidden = false;
        } else {
          canvas.style.display = '';
          if (emptyMsg) emptyMsg.hidden = true;
        }
      }
    }
  },

  /**
   * Prepare Chart.js data object for a given chart configuration.
   * Applies data aggregation for line and bar charts when data exceeds MAX_CHART_POINTS.
   * @param {Object} chartConfig - Chart configuration from CHART_CONFIGS
   * @param {FinancialRecord[]} data - Financial records
   * @returns {Object} Chart.js data object with labels and datasets
   */
  _prepareChartData(chartConfig, data) {
    if (!data || data.length === 0) {
      return { labels: [], datasets: [] };
    }

    let chartData;
    switch (chartConfig.id) {
      case 'fluxo-caixa':
        chartData = this._prepareFluxoCaixa(data);
        break;
      case 'receitas-chart':
        chartData = this._prepareReceitasChart(data);
        break;
      case 'despesas-chart':
        chartData = this._prepareDespesasChart(data);
        break;
      case 'recebimentos-proj':
        chartData = this._prepareRecebimentosProj(data);
        break;
      case 'evolucao-mensal':
        chartData = this._prepareEvolucaoMensal(data);
        break;
      case 'top-clientes':
        chartData = this._prepareTopClientes(data);
        break;
      case 'tipo-pagamento':
        chartData = this._prepareTipoPagamentoData(data);
        break;
      default:
        return { labels: [], datasets: [] };
    }

    // Apply aggregation for large datasets (line and bar charts only)
    return this._applyAggregationIfNeeded(chartData, chartConfig.type);
  },

  /**
   * Group records by month key (YYYY-MM) and return sorted month keys.
   * @param {FinancialRecord[]} data - Financial records
   * @returns {{monthKeys: string[], monthGroups: Map<string, FinancialRecord[]>}}
   * @private
   */
  _groupByMonth(data) {
    const monthGroups = new Map();
    for (const record of data) {
      const key = `${record.data.getFullYear()}-${String(record.data.getMonth() + 1).padStart(2, '0')}`;
      if (!monthGroups.has(key)) {
        monthGroups.set(key, []);
      }
      monthGroups.get(key).push(record);
    }
    const monthKeys = [...monthGroups.keys()].sort();
    return { monthKeys, monthGroups };
  },

  /**
   * Format a month key (YYYY-MM) to a readable label (e.g., "Jan/2024").
   * @param {string} monthKey - Month key in YYYY-MM format
   * @returns {string} Formatted month label
   * @private
   */
  _formatMonthLabel(monthKey) {
    const [year, month] = monthKey.split('-');
    const monthNames = ['Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez'];
    return `${monthNames[parseInt(month) - 1]}/${year}`;
  },

  /**
   * Fluxo de Caixa Mensal: bar chart of net cash flow per month.
   * Positive bars in green, negative bars in red.
   * @private
   */
  _prepareFluxoCaixa(data) {
    const { monthKeys, monthGroups } = this._groupByMonth(data);
    const values = monthKeys.map(key => {
      const records = monthGroups.get(key);
      const receitas = records.filter(r => r.tipo === 'receita').reduce((s, r) => s + r.valor, 0);
      const despesas = records.filter(r => r.tipo === 'despesa').reduce((s, r) => s + r.valor, 0);
      return receitas - despesas;
    });

    const backgroundColors = values.map((_, i) => this._monthPalette[i % 12]);

    return {
      labels: monthKeys.map(k => this._formatMonthLabel(k)),
      datasets: [{
        label: 'Fluxo de Caixa',
        data: values,
        backgroundColor: backgroundColors,
        borderColor: backgroundColors,
        borderWidth: 1
      }]
    };
  },

  /**
   * Receitas vs Despesas: grouped bar chart with receitas and despesas per month.
   * @private
   */
  _prepareReceitasDespesas(data) {
    const { monthKeys, monthGroups } = this._groupByMonth(data);
    const receitas = monthKeys.map(key => {
      return monthGroups.get(key).filter(r => r.tipo === 'receita').reduce((s, r) => s + r.valor, 0);
    });
    const despesas = monthKeys.map(key => {
      return monthGroups.get(key).filter(r => r.tipo === 'despesa').reduce((s, r) => s + r.valor, 0);
    });

    return {
      labels: monthKeys.map(k => this._formatMonthLabel(k)),
      datasets: [
        {
          label: 'Receitas',
          data: receitas,
          backgroundColor: this._positiveColor,
          borderColor: this._positiveColor,
          borderWidth: 1
        },
        {
          label: 'Despesas',
          data: despesas,
          backgroundColor: this._negativeColor,
          borderColor: this._negativeColor,
          borderWidth: 1
        }
      ]
    };
  },

  /**
   * Recebimentos por Projeto: horizontal bar of receitas grouped by project.
   * @private
   */
  _prepareRecebimentosProj(data) {
    const projectTotals = {};
    for (const record of data) {
      // Soma TODOS os valores de receita vinculados a cada projeto
      if (record.tipo === 'receita' && record.projeto) {
        projectTotals[record.projeto] = (projectTotals[record.projeto] || 0) + record.valor;
      }
    }

    // Sort by value descending, take top 10
    const sorted = Object.entries(projectTotals)
      .sort((a, b) => b[1] - a[1])
      .slice(0, 15);

    const labels = sorted.map(([name]) => name);
    const values = sorted.map(([, val]) => val);
    const colors = labels.map((_, i) => this._colors[i % this._colors.length]);

    return {
      labels,
      datasets: [{
        label: 'Total Recebido',
        data: values,
        backgroundColor: colors,
        borderColor: colors,
        borderWidth: 1
      }]
    };
  },

  /**
   * Despesas por Categoria: doughnut chart of despesas grouped by category.
   * @private
   */
  _prepareDespesasCat(data) {
    const categoryTotals = {};
    for (const record of data) {
      if (record.tipo === 'despesa' && record.categoria) {
        categoryTotals[record.categoria] = (categoryTotals[record.categoria] || 0) + record.valor;
      }
    }

    // Sort by value descending
    const sorted = Object.entries(categoryTotals)
      .sort((a, b) => b[1] - a[1]);

    const labels = sorted.map(([name]) => name);
    const values = sorted.map(([, val]) => val);
    const colors = labels.map((_, i) => this._colors[i % this._colors.length]);

    return {
      labels,
      datasets: [{
        label: 'Despesas',
        data: values,
        backgroundColor: colors,
        borderColor: '#ffffff',
        borderWidth: 2
      }]
    };
  },

  /**
   * Classificação Geral: doughnut of tipoClassif (Receitas, Despesas, Custos, Saldo Inicial, etc.)
   * @private
   */
  _prepareClassificacaoGeral(data) {
    const totals = {};
    for (const record of data) {
      if (record.tipoClassif) {
        totals[record.tipoClassif] = (totals[record.tipoClassif] || 0) + record.valor;
      }
    }
    const sorted = Object.entries(totals).sort((a, b) => b[1] - a[1]);
    const labels = sorted.map(([name]) => name);
    const values = sorted.map(([, val]) => val);
    const colors = labels.map((_, i) => this._colors[i % this._colors.length]);
    return {
      labels,
      datasets: [{ label: 'Classificação', data: values, backgroundColor: colors, borderColor: '#ffffff', borderWidth: 2 }]
    };
  },

  /**
   * Receitas chart: doughnut showing only receita records grouped by nivel1 (2ª Categoria).
   * @private
   */
  _prepareReceitasChart(data) {
    // Only records from the "+" column (tipo === 'receita')
    const receitas = data.filter(r => r.tipo === 'receita');
    const totals = {};
    for (const record of receitas) {
      const cat = record.nivel1 || record.nivel2 || record.tipoClassif || 'Outros';
      totals[cat] = (totals[cat] || 0) + record.valor;
    }
    // Remove any category names that clearly indicate expenses (safety filter)
    const sorted = Object.entries(totals)
      .filter(([name]) => !/despesa|custo|sa[ií]da|pagamento|m[aã]o de obra/i.test(name))
      .sort((a, b) => b[1] - a[1]).slice(0, 10);
    const labels = sorted.map(([name]) => name);
    const values = sorted.map(([, val]) => val);
    const colors = labels.map((_, i) => this._colors[i % this._colors.length]);
    return {
      labels,
      datasets: [{ label: 'Receitas', data: values, backgroundColor: colors, borderColor: '#ffffff', borderWidth: 2 }]
    };
  },

  /**
   * Despesas chart: doughnut showing only despesa records grouped by nivel1 (2ª Categoria).
   * @private
   */
  _prepareDespesasChart(data) {
    // Only records from the "-" column (tipo === 'despesa')
    const despesas = data.filter(r => r.tipo === 'despesa');
    const totals = {};
    for (const record of despesas) {
      const cat = record.nivel1 || record.nivel2 || record.tipoClassif || 'Outros';
      totals[cat] = (totals[cat] || 0) + record.valor;
    }
    // Remove any category names that clearly indicate revenues (safety filter)
    const sorted = Object.entries(totals)
      .filter(([name]) => !/receita|lucro|entrada|cr[eé]dito/i.test(name))
      .sort((a, b) => b[1] - a[1]).slice(0, 10);
    const labels = sorted.map(([name]) => name);
    const values = sorted.map(([, val]) => val);
    const colors = labels.map((_, i) => this._colors[i % this._colors.length]);
    return {
      labels,
      datasets: [{ label: 'Despesas', data: values, backgroundColor: colors, borderColor: '#ffffff', borderWidth: 2 }]
    };
  },

  /**
   * Nível 2 chart: doughnut of nivel1 (1ª Categoria) filtered by dropdown (receita/despesa/all).
   * @private
   */
  _prepareNivel2Chart(data) {
    const filterEl = document.getElementById('nivel2-filter');
    const filterValue = filterEl ? filterEl.value : 'all';
    
    let filtered = data;
    if (filterValue === 'receita') {
      filtered = data.filter(r => r.tipo === 'receita');
    } else if (filterValue === 'despesa') {
      filtered = data.filter(r => r.tipo === 'despesa');
    }

    const totals = {};
    for (const record of filtered) {
      if (record.nivel1) {
        totals[record.nivel1] = (totals[record.nivel1] || 0) + record.valor;
      }
    }
    const sorted = Object.entries(totals).sort((a, b) => b[1] - a[1]);
    const labels = sorted.map(([name]) => name);
    const values = sorted.map(([, val]) => val);
    const colors = labels.map((_, i) => this._colors[i % this._colors.length]);
    return {
      labels,
      datasets: [{ label: '1ª Categoria', data: values, backgroundColor: colors, borderColor: '#ffffff', borderWidth: 2 }]
    };
  },

  /**
   * Nível 3 chart: horizontal bar of nivel2 (2ª Categoria) filtered by dropdown (receita/despesa/all).
   * @private
   */
  _prepareNivel3Chart(data) {
    const filterEl = document.getElementById('nivel3-filter');
    const filterValue = filterEl ? filterEl.value : 'all';
    
    let filtered = data;
    if (filterValue === 'receita') {
      filtered = data.filter(r => r.tipo === 'receita');
    } else if (filterValue === 'despesa') {
      filtered = data.filter(r => r.tipo === 'despesa');
    }

    const totals = {};
    for (const record of filtered) {
      if (record.nivel2) {
        totals[record.nivel2] = (totals[record.nivel2] || 0) + record.valor;
      }
    }
    const sorted = Object.entries(totals).sort((a, b) => b[1] - a[1]).slice(0, 15);
    const labels = sorted.map(([name]) => name);
    const values = sorted.map(([, val]) => val);
    const colors = labels.map((_, i) => this._colors[i % this._colors.length]);
    return {
      labels,
      datasets: [{ label: '2ª Categoria', data: values, backgroundColor: colors, borderColor: colors, borderWidth: 1 }]
    };
  },

  /**
   * Evolução Financeira Mensal: line chart of monthly receitas and despesas totals.
   * @private
   */
  _prepareEvolucaoMensal(data) {
    const { monthKeys, monthGroups } = this._groupByMonth(data);
    const receitas = monthKeys.map(key => {
      return monthGroups.get(key).filter(r => r.tipo === 'receita').reduce((s, r) => s + r.valor, 0);
    });
    const despesas = monthKeys.map(key => {
      return monthGroups.get(key).filter(r => r.tipo === 'despesa').reduce((s, r) => s + r.valor, 0);
    });

    return {
      labels: monthKeys.map(k => this._formatMonthLabel(k)),
      datasets: [
        {
          label: 'Receitas',
          data: receitas,
          borderColor: this._positiveColor,
          backgroundColor: 'rgba(13, 122, 63, 0.1)',
          fill: true,
          tension: 0.3,
          pointRadius: 4,
          pointHoverRadius: 6
        },
        {
          label: 'Despesas',
          data: despesas,
          borderColor: this._negativeColor,
          backgroundColor: 'rgba(196, 30, 58, 0.1)',
          fill: true,
          tension: 0.3,
          pointRadius: 4,
          pointHoverRadius: 6
        }
      ]
    };
  },

  /**
   * Status Financeiro: doughnut chart of records grouped by status.
   * @private
   */
  _prepareStatusFinanceiro(data) {
    const statusTotals = {};
    for (const record of data) {
      const status = record.status || 'pendente';
      statusTotals[status] = (statusTotals[status] || 0) + 1;
    }

    const statusColors = {
      'recebido': '#0d7a3f',
      'pago': '#1d4ed8',
      'pendente': '#b45309',
      'vencido': '#c41e3a'
    };

    const labels = Object.keys(statusTotals);
    const values = Object.values(statusTotals);
    const colors = labels.map(s => statusColors[s] || this._colors[0]);

    return {
      labels,
      datasets: [{
        label: 'Status',
        data: values,
        backgroundColor: colors,
        borderColor: '#ffffff',
        borderWidth: 2
      }]
    };
  },

  /**
   * Top Clientes: horizontal bar of top clients by revenue.
   * @private
   */
  _prepareTopClientes(data) {
    const clientTotals = {};
    for (const record of data) {
      // Soma TODOS os valores de receita vinculados a cada De/Para (cliente)
      if (record.tipo === 'receita' && record.cliente) {
        clientTotals[record.cliente] = (clientTotals[record.cliente] || 0) + record.valor;
      }
    }

    // Sort by value descending, take top 15
    const sorted = Object.entries(clientTotals)
      .sort((a, b) => b[1] - a[1])
      .slice(0, 15);

    const labels = sorted.map(([name]) => name);
    const values = sorted.map(([, val]) => val);
    const colors = labels.map((_, i) => this._colors[i % this._colors.length]);

    return {
      labels,
      datasets: [{
        label: 'Total Recebido',
        data: values,
        backgroundColor: colors,
        borderColor: colors,
        borderWidth: 1
      }]
    };
  },

  /**
   * Tipo de Pagamento: doughnut chart grouped by tipoPagamento field.
   * @param {FinancialRecord[]} data - Financial records
   * @returns {Object|null} Chart.js data object or null if no valid data
   * @private
   */
  _prepareTipoPagamentoData(data) {
    const totals = {};
    for (const record of data) {
      if (record.tipoPagamento && record.tipoPagamento.trim() !== '') {
        const key = record.tipoPagamento.trim();
        totals[key] = (totals[key] || 0) + Math.abs(record.valor);
      }
    }

    const sorted = Object.entries(totals).sort((a, b) => b[1] - a[1]);

    if (sorted.length === 0) {
      return { labels: [], datasets: [] };
    }

    const labels = sorted.map(([name]) => name);
    const values = sorted.map(([, val]) => val);
    const colors = labels.map((_, i) => this._colors[i % this._colors.length]);

    return {
      labels,
      datasets: [{
        label: 'Tipo de Pagamento',
        data: values,
        backgroundColor: colors,
        borderColor: colors,
        borderWidth: 1
      }]
    };
  },

  /**
   * Inadimplência: bar chart of overdue (vencido) amounts by month.
   * @private
   */
  _prepareInadimplencia(data) {
    const overdueRecords = data.filter(r => r.status === 'Pendente');
    const { monthKeys, monthGroups } = this._groupByMonth(overdueRecords);

    const values = monthKeys.map(key => {
      return monthGroups.get(key).reduce((s, r) => s + r.valor, 0);
    });

    return {
      labels: monthKeys.map(k => this._formatMonthLabel(k)),
      datasets: [{
        label: 'Inadimplência',
        data: values,
        backgroundColor: this._negativeColor,
        borderColor: this._negativeColor,
        borderWidth: 1
      }]
    };
  },

  /**
   * Curva de Recebimentos: cumulative line chart of received amounts over time.
   * @private
   */
  _prepareCurvaRecebimentos(data) {
    const receivedRecords = data.filter(r => r.tipo === 'receita' && r.status === 'Pago');
    const { monthKeys, monthGroups } = this._groupByMonth(receivedRecords);

    let cumulative = 0;
    const values = monthKeys.map(key => {
      const monthTotal = monthGroups.get(key).reduce((s, r) => s + r.valor, 0);
      cumulative += monthTotal;
      return cumulative;
    });

    return {
      labels: monthKeys.map(k => this._formatMonthLabel(k)),
      datasets: [{
        label: 'Recebimentos Acumulados',
        data: values,
        borderColor: this._colors[0],
        backgroundColor: 'rgba(29, 78, 216, 0.1)',
        fill: true,
        tension: 0.3,
        pointRadius: 4,
        pointHoverRadius: 6
      }]
    };
  },

  /**
   * Comparativo Mensal: grouped bar comparing current period vs previous period.
   * Splits data into two halves and compares month-by-month.
   * @private
   */
  _prepareComparativoMensal(data) {
    const { monthKeys, monthGroups } = this._groupByMonth(data);

    if (monthKeys.length < 2) {
      // Not enough months for comparison
      const values = monthKeys.map(key => {
        return monthGroups.get(key).reduce((s, r) => s + r.valor, 0);
      });
      return {
        labels: monthKeys.map(k => this._formatMonthLabel(k)),
        datasets: [{
          label: 'Período Atual',
          data: values,
          backgroundColor: this._colors[0],
          borderColor: this._colors[0],
          borderWidth: 1
        }]
      };
    }

    // Split months into two halves: previous period and current period
    const midpoint = Math.ceil(monthKeys.length / 2);
    const previousKeys = monthKeys.slice(0, midpoint);
    const currentKeys = monthKeys.slice(midpoint);

    // Use the longer set for labels (month names only, no year for comparison)
    const maxLen = Math.max(previousKeys.length, currentKeys.length);
    const labels = [];
    const previousValues = [];
    const currentValues = [];

    for (let i = 0; i < maxLen; i++) {
      if (i < currentKeys.length) {
        labels.push(this._formatMonthLabel(currentKeys[i]));
        currentValues.push(monthGroups.get(currentKeys[i]).reduce((s, r) => s + r.valor, 0));
      } else {
        labels.push('');
        currentValues.push(0);
      }

      if (i < previousKeys.length) {
        previousValues.push(monthGroups.get(previousKeys[i]).reduce((s, r) => s + r.valor, 0));
      } else {
        previousValues.push(0);
      }
    }

    return {
      labels,
      datasets: [
        {
          label: 'Período Anterior',
          data: previousValues,
          backgroundColor: this._colors[4],
          borderColor: this._colors[4],
          borderWidth: 1
        },
        {
          label: 'Período Atual',
          data: currentValues,
          backgroundColor: this._colors[0],
          borderColor: this._colors[0],
          borderWidth: 1
        }
      ]
    };
  },

  /**
   * Get Chart.js options for a given chart configuration.
   * Configures scales, animations, plugins, and layout based on chart type.
   * @param {Object} chartConfig - Chart configuration from CHART_CONFIGS
   * @returns {Object} Chart.js options object
   */
  _getChartOptions(chartConfig) {
    const baseOptions = {
      responsive: true,
      maintainAspectRatio: false,
      animation: {
        duration: PERFORMANCE.ANIMATION_DURATION
      },
      plugins: {
        legend: {
          display: chartConfig.grouped || chartConfig.type === 'doughnut' || chartConfig.type === 'line',
          position: 'top',
          labels: {
            usePointStyle: true,
            padding: 12,
            font: { size: 12 }
          }
        },
        tooltip: {
          enabled: true,
          backgroundColor: 'rgba(0, 0, 0, 0.8)',
          titleFont: { size: 13 },
          bodyFont: { size: 12 },
          padding: 10,
          cornerRadius: 6,
          callbacks: {
            label: function(context) {
              const label = context.dataset.label || '';
              // Use context.raw for the actual value (works for both horizontal and vertical charts)
              const value = context.raw;
              const formatted = KPICalculator.formatBRL(typeof value === 'number' ? value : 0);
              // Calculate percentage for doughnut charts
              if (chartConfig.type === 'doughnut') {
                const total = context.dataset.data.reduce((a, b) => a + b, 0);
                const pct = total > 0 ? ((value / total) * 100).toFixed(1) : '0.0';
                return `${label}: ${formatted} (${pct}%)`;
              }
              return `${label}: ${formatted}`;
            }
          }
        }
      }
    };

    // Doughnut-specific options (no scales)
    if (chartConfig.type === 'doughnut') {
      baseOptions.cutout = '55%';
      return baseOptions;
    }

    // Scale options for bar and line charts
    baseOptions.scales = {};

    if (chartConfig.horizontal) {
      baseOptions.indexAxis = 'y';
      baseOptions.scales.x = {
        beginAtZero: true,
        ticks: {
          callback: function(value) {
            return KPICalculator.formatBRL(value);
          },
          font: { size: 11 }
        },
        grid: { color: 'rgba(0, 0, 0, 0.05)' }
      };
      baseOptions.scales.y = {
        ticks: { font: { size: 11 }, color: '#1d4ed8' },
        grid: { display: false }
      };
      // Add onHover to show pointer cursor on Y-axis labels
      baseOptions.onHover = (event, elements, chart) => {
        const nativeEvent = event.native || event;
        const rect = chart.canvas.getBoundingClientRect();
        const x = nativeEvent.clientX - rect.left;
        if (x < chart.chartArea.left + 5) {
          chart.canvas.style.cursor = 'pointer';
        } else if (elements && elements.length > 0) {
          chart.canvas.style.cursor = 'pointer';
        } else {
          chart.canvas.style.cursor = 'default';
        }
      };
    } else {
      baseOptions.scales.x = {
        ticks: { font: { size: 11 } },
        grid: { display: false }
      };
      baseOptions.scales.y = {
        beginAtZero: true,
        ticks: {
          callback: function(value) {
            return KPICalculator.formatBRL(value);
          },
          font: { size: 11 }
        },
        grid: { color: 'rgba(0, 0, 0, 0.05)' }
      };
    }

    return baseOptions;
  },

  /**
   * Check if data points exceed the MAX_CHART_POINTS threshold and need aggregation.
   * @param {Array} dataPoints - Array of data points to check
   * @returns {boolean} True if aggregation is needed
   */
  _shouldAggregate(dataPoints) {
    return Array.isArray(dataPoints) && dataPoints.length > PERFORMANCE.MAX_CHART_POINTS;
  },

  /**
   * Aggregate data points when count exceeds threshold.
   * Groups adjacent points into buckets, summing their values.
   * Preserves total sum: sum of aggregated === sum of original.
   * @param {number[]} dataPoints - Raw chart data values
   * @param {number} [maxPoints] - Maximum points to render (defaults to PERFORMANCE.MAX_CHART_POINTS)
   * @returns {number[]} Aggregated data values
   */
  aggregateData(dataPoints, maxPoints) {
    const max = maxPoints || PERFORMANCE.MAX_CHART_POINTS;
    if (!dataPoints || dataPoints.length <= max) return dataPoints;

    const groupSize = Math.ceil(dataPoints.length / max);
    const aggregated = [];

    for (let i = 0; i < dataPoints.length; i += groupSize) {
      const group = dataPoints.slice(i, i + groupSize);
      const sum = group.reduce((s, p) => s + (typeof p === 'number' ? p : (p.value || 0)), 0);
      aggregated.push(sum);
    }

    return aggregated;
  },

  /**
   * Aggregate labels to match aggregated data points.
   * Takes the first label of each group as the representative label.
   * @param {string[]} labels - Original labels array
   * @param {number} [maxPoints] - Maximum points (defaults to PERFORMANCE.MAX_CHART_POINTS)
   * @returns {string[]} Aggregated labels
   */
  _aggregateLabels(labels, maxPoints) {
    const max = maxPoints || PERFORMANCE.MAX_CHART_POINTS;
    if (!labels || labels.length <= max) return labels;

    const groupSize = Math.ceil(labels.length / max);
    const aggregated = [];

    for (let i = 0; i < labels.length; i += groupSize) {
      const groupEnd = Math.min(i + groupSize - 1, labels.length - 1);
      // Use range label if group spans multiple items
      if (i === groupEnd) {
        aggregated.push(labels[i]);
      } else {
        aggregated.push(`${labels[i]} - ${labels[groupEnd]}`);
      }
    }

    return aggregated;
  },

  /**
   * Apply aggregation to a prepared chart data object if any dataset exceeds the threshold.
   * Aggregates all datasets and labels consistently.
   * Only applies to bar and line charts (not doughnut).
   * @param {Object} chartData - Chart.js data object {labels, datasets}
   * @param {string} chartType - Chart type ('bar', 'line', 'doughnut')
   * @returns {Object} Possibly aggregated chart data
   */
  _applyAggregationIfNeeded(chartData, chartType) {
    // Don't aggregate doughnut charts - they typically have few categories
    if (chartType === 'doughnut') return chartData;
    if (!chartData || !chartData.datasets || chartData.datasets.length === 0) return chartData;

    // Check if any dataset exceeds the threshold
    const needsAggregation = chartData.datasets.some(ds => this._shouldAggregate(ds.data));
    if (!needsAggregation) return chartData;

    // Aggregate labels
    const aggregatedLabels = this._aggregateLabels(chartData.labels);

    // Aggregate each dataset's data array
    const aggregatedDatasets = chartData.datasets.map(ds => ({
      ...ds,
      data: this.aggregateData(ds.data)
    }));

    return {
      labels: aggregatedLabels,
      datasets: aggregatedDatasets
    };
  },

  /**
   * Derive a filter from a chart click event.
   * @param {string} chartId - Chart identifier
   * @param {Object} element - Clicked chart element info
   * @returns {FilterCriteria} Filter to apply
   */
  getFilterFromClick(chartId, element) {
    if (!element || element.index == null) return {};

    const chart = this._instances[chartId];
    if (!chart) return {};

    const label = chart.data.labels[element.index];
    if (!label) return {};

    switch (chartId) {
      case 'recebimentos-proj':
        return { projeto: label };
      case 'receitas-nivel2':
      case 'despesas-nivel2':
        return { nivel2: label };
      case 'receitas-nivel3':
      case 'despesas-nivel3':
        return { nivel3: label };
      case 'top-clientes':
        return { cliente: label };
      case 'status-financeiro':
        return { status: label };
      case 'tipo-pagamento':
        return { tipoPagamento: label };
      default:
        return {};
    }
  },

  /**
   * Export a specific chart as PNG image.
   * @param {string} chartId - Chart identifier
   * @returns {Promise<Blob>} PNG image blob
   */
  exportChartImage(chartId) {
    return new Promise((resolve, reject) => {
      const chart = this._instances[chartId];
      if (!chart) {
        reject(new Error(`Chart "${chartId}" not found`));
        return;
      }

      const canvas = chart.canvas;
      canvas.toBlob((blob) => {
        if (blob) {
          resolve(blob);
        } else {
          reject(new Error('Failed to export chart as image'));
        }
      }, 'image/png');
    });
  }
};

// Wire ChartEngine to EventBus events
// Note: ChartEngine.initCharts is called by UIController's lazy rendering strategy
// on data:loaded (priority section), so we skip it here to avoid double initialization.

EventBus.on('filters:changed', (payload) => {
  // Only update charts if the chart section has been rendered
  if (typeof UIController !== 'undefined' && UIController._sectionRendered &&
      UIController._sectionRendered.has('section-financeiro')) {
    ChartEngine.updateCharts(payload.filteredData);
  }
});

// Wire chart:clicked event to open detail popup for project/client charts
EventBus.on('chart:clicked', (payload) => {
  if (!payload || !payload.filter) return;
  
  const chartId = payload.chartId;
  
  // For project and client charts, open a popup with filtered records
  if (chartId === 'recebimentos-proj' || chartId === 'top-clientes' || chartId === 'tipo-pagamento') {
    const data = UIController._currentData || DataLayer.getData();
    let filtered = data;
    let title = '';
    
    if (payload.filter.projeto) {
      filtered = data.filter(r => r.projeto === payload.filter.projeto);
      title = 'Lançamentos: ' + payload.filter.projeto;
    } else if (payload.filter.cliente) {
      filtered = data.filter(r => r.cliente === payload.filter.cliente);
      title = 'Lançamentos: ' + payload.filter.cliente;
    } else if (payload.filter.tipoPagamento) {
      filtered = data.filter(r => r.tipoPagamento === payload.filter.tipoPagamento);
      title = 'Lançamentos: ' + payload.filter.tipoPagamento;
    }
    
    if (window.KPIDetailModal) {
      KPIDetailModal._kpiKey = 'custom';
      KPIDetailModal._page = 1;
      KPIDetailModal._sortColumn = 'data';
      KPIDetailModal._sortDirection = 'asc';
      KPIDetailModal._records = filtered;
      const titleEl = document.getElementById('kpi-detail-title');
      if (titleEl) titleEl.textContent = title;
      KPIDetailModal._renderSummary();
      KPIDetailModal._renderTable();
      const modal = document.getElementById('kpi-detail-modal');
      if (modal) modal.removeAttribute('hidden');
      document.body.style.overflow = 'hidden';
    }
    return;
  }
  
  // For other charts, apply as filter
  FilterEngine.applyFiltersDebounced(payload.filter);
});

// ============================================================
// SECTION: Table Module
// ============================================================

/** @type {Object} TableModule - Sortable, searchable, paginated tables */
const TableModule = {
  /** @type {number|null} Internal timer reference for search debounce */
  _searchDebounceTimer: null,

  /**
   * Sort records by a column.
   * Uses locale-aware comparison for strings (pt-BR), numeric comparison for numbers,
   * and chronological comparison for dates.
   * Does not mutate the input array.
   *
   * @param {FinancialRecord[]} data - Records to sort
   * @param {string} column - Column name
   * @param {'asc'|'desc'} direction - Sort direction
   * @returns {FinancialRecord[]} New sorted array
   */
  sortByColumn(data, column, direction) {
    if (!data || !Array.isArray(data)) return [];
    if (!column) return [...data];

    const dir = direction === 'desc' ? -1 : 1;

    return [...data].sort((a, b) => {
      const valA = a[column];
      const valB = b[column];

      // Handle null/undefined values - push them to the end
      if (valA == null && valB == null) return 0;
      if (valA == null) return 1;
      if (valB == null) return -1;

      // Numeric comparison for 'valor' column
      if (column === 'valor') {
        return (Number(valA) - Number(valB)) * dir;
      }

      // Date comparison for 'data' column
      if (column === 'data') {
        const dateA = valA instanceof Date ? valA.getTime() : new Date(valA).getTime();
        const dateB = valB instanceof Date ? valB.getTime() : new Date(valB).getTime();
        return (dateA - dateB) * dir;
      }

      // String comparison with locale-aware sorting (pt-BR)
      const strA = String(valA);
      const strB = String(valB);
      return strA.localeCompare(strB, 'pt-BR') * dir;
    });
  },

  /**
   * Search records with case-insensitive partial matching across all visible columns.
   * If query is empty, returns all data.
   *
   * @param {FinancialRecord[]} data - Records to search
   * @param {string} query - Search term
   * @returns {FinancialRecord[]} Matching records
   */
  searchRecords(data, query) {
    if (!data || !Array.isArray(data)) return [];
    if (!query || query.trim() === '') return data;

    const lowerQuery = query.toLowerCase();
    const searchColumns = ['projeto', 'cliente', 'categoria', 'valor', 'tipo', 'data', 'status', 'centroCusto', 'responsavel'];

    return data.filter(record => {
      return searchColumns.some(col => {
        const value = record[col];
        if (value == null) return false;

        // Format date for search comparison
        if (col === 'data' && value instanceof Date) {
          const formatted = this.formatDate(value);
          return formatted.toLowerCase().includes(lowerQuery);
        }

        // Convert to string for comparison
        return String(value).toLowerCase().includes(lowerQuery);
      });
    });
  },

  /**
   * Paginate records with PAGE_SIZE = PERFORMANCE.PAGE_SIZE (20).
   * Page is 1-based and clamped to valid range.
   *
   * @param {FinancialRecord[]} data - Records to paginate
   * @param {number} page - Page number (1-based)
   * @returns {{rows: FinancialRecord[], currentPage: number, totalPages: number}}
   */
  paginate(data, page) {
    if (!data || !Array.isArray(data) || data.length === 0) {
      return { rows: [], currentPage: 1, totalPages: 0 };
    }

    const pageSize = PERFORMANCE.PAGE_SIZE;
    const totalPages = Math.ceil(data.length / pageSize);

    // Clamp page to valid range
    let currentPage = Math.max(1, Math.min(page || 1, totalPages));

    const startIndex = (currentPage - 1) * pageSize;
    const endIndex = Math.min(startIndex + pageSize, data.length);
    const rows = data.slice(startIndex, endIndex);

    return { rows, currentPage, totalPages };
  },

  /**
   * Format a Date as DD/MM/YYYY.
   * Pads day and month with leading zeros.
   *
   * @param {Date|string} date - Date to format
   * @returns {string} Formatted date string (e.g., "15/01/2024")
   */
  formatDate(date) {
    if (!date) return '';

    const d = date instanceof Date ? date : new Date(date);
    if (isNaN(d.getTime())) return '';

    const day = String(d.getDate()).padStart(2, '0');
    const month = String(d.getMonth() + 1).padStart(2, '0');
    const year = d.getFullYear();

    return `${day}/${month}/${year}`;
  },

  /**
   * Debounced search that clears previous timer and sets a new one.
   * When timer fires, calls callback with searchRecords result.
   *
   * @param {FinancialRecord[]} data - Records to search
   * @param {string} query - Search term
   * @param {function} callback - Callback receiving filtered results
   */
  searchDebounced(data, query, callback) {
    if (this._searchDebounceTimer) {
      clearTimeout(this._searchDebounceTimer);
    }

    this._searchDebounceTimer = setTimeout(() => {
      const results = this.searchRecords(data, query);
      if (typeof callback === 'function') {
        callback(results);
      }
    }, PERFORMANCE.SEARCH_DEBOUNCE);
  }
};

// ============================================================
// SECTION: Insights Engine
// ============================================================

/** @type {Object} InsightsEngine - Automatic financial analysis */
const InsightsEngine = {
  /**
   * Generate all insights from financial data.
   * Computes monthly changes, rankings, overdue alerts, top clients/categories,
   * trend direction, and returns a complete InsightResult object.
   * @param {FinancialRecord[]} data - Financial records
   * @returns {InsightResult} All computed insights
   */
  generateInsights(data) {
    if (!data || !Array.isArray(data) || data.length === 0) {
      return {
        expenseIncreases: [],
        revenueDrops: [],
        topClients: [],
        topCategories: [],
        monthlyVariations: [],
        alerts: [],
        trendDirection: '→',
        projectRanking: [],
        clientRanking: [],
        categoryRanking: [],
        forecast: null
      };
    }

    const monthlyChanges = this.calculateMonthlyChanges(data);
    const rankings = this.generateRankings(data);
    const alerts = this.detectOverdueAlerts(data);
    const topClients = this.getTopClients(data, 5);
    const topCategories = this.getTopExpenseCategories(data, 5);

    // Compute monthly totals for trend direction
    const monthlyTotals = this._getMonthlyNetTotals(data);
    const trendDirection = this.calculateTrendDirection(monthlyTotals);

    // Generate forecast (returns null if fewer than 3 months)
    const forecast = this.generateForecast(data);

    // Detect critical conditions from monthly changes
    const criticalConditions = this.detectCriticalConditions(monthlyChanges);

    // Identify expense increases >10% and revenue drops >5%
    const expenseIncreases = monthlyChanges.expenseChanges.filter(
      c => c.change !== null && c.change > 10
    );
    const revenueDrops = monthlyChanges.revenueChanges.filter(
      c => c.change !== null && c.change < -5
    );

    // Combine all monthly variations
    const monthlyVariations = [
      ...monthlyChanges.expenseChanges,
      ...monthlyChanges.revenueChanges
    ];

    return {
      expenseIncreases,
      revenueDrops,
      topClients,
      topCategories,
      monthlyVariations,
      alerts: [...alerts, ...criticalConditions],
      trendDirection,
      projectRanking: rankings.projects,
      clientRanking: rankings.clients,
      categoryRanking: rankings.categories,
      forecast
    };
  },

  /**
   * Calculate month-over-month changes for expenses and revenue.
   * Groups data by month (YYYY-MM), computes totals per month,
   * and calculates percentage change from previous month.
   * @param {FinancialRecord[]} data - Financial records
   * @returns {{expenseChanges: Array<{month: string, value: number, previousValue: number|null, change: number|null}>, revenueChanges: Array<{month: string, value: number, previousValue: number|null, change: number|null}>}}
   */
  calculateMonthlyChanges(data) {
    if (!data || !Array.isArray(data) || data.length === 0) {
      return { expenseChanges: [], revenueChanges: [] };
    }

    // Group by month (YYYY-MM)
    const expensesByMonth = {};
    const revenueByMonth = {};

    for (const record of data) {
      if (!record.data || !(record.data instanceof Date) || isNaN(record.data.getTime())) continue;

      const year = record.data.getFullYear();
      const month = String(record.data.getMonth() + 1).padStart(2, '0');
      const monthKey = `${year}-${month}`;

      if (record.tipo === 'despesa') {
        expensesByMonth[monthKey] = (expensesByMonth[monthKey] || 0) + record.valor;
      } else if (record.tipo === 'receita') {
        revenueByMonth[monthKey] = (revenueByMonth[monthKey] || 0) + record.valor;
      }
    }

    // Get all months sorted chronologically
    const allMonths = [...new Set([
      ...Object.keys(expensesByMonth),
      ...Object.keys(revenueByMonth)
    ])].sort();

    // Calculate expense changes
    const expenseChanges = [];
    for (let i = 0; i < allMonths.length; i++) {
      const month = allMonths[i];
      const value = expensesByMonth[month] || 0;
      const previousValue = i > 0 ? (expensesByMonth[allMonths[i - 1]] || 0) : null;
      let change = null;

      if (previousValue !== null && previousValue !== 0) {
        change = ((value - previousValue) / Math.abs(previousValue)) * 100;
      } else if (previousValue === 0 && value > 0) {
        change = 100; // From zero to something is 100% increase
      } else if (previousValue === null) {
        change = null; // No previous month to compare
      }

      expenseChanges.push({ month, value, previousValue, change });
    }

    // Calculate revenue changes
    const revenueChanges = [];
    for (let i = 0; i < allMonths.length; i++) {
      const month = allMonths[i];
      const value = revenueByMonth[month] || 0;
      const previousValue = i > 0 ? (revenueByMonth[allMonths[i - 1]] || 0) : null;
      let change = null;

      if (previousValue !== null && previousValue !== 0) {
        change = ((value - previousValue) / Math.abs(previousValue)) * 100;
      } else if (previousValue === 0 && value > 0) {
        change = 100;
      } else if (previousValue === null) {
        change = null;
      }

      revenueChanges.push({ month, value, previousValue, change });
    }

    return { expenseChanges, revenueChanges };
  },

  /**
   * Generate rankings (top 10) for projects, clients, categories.
   * Ranked by net value (sum of receita - sum of despesa) in descending order.
   * @param {FinancialRecord[]} data - Financial records
   * @returns {{projects: Array<{name: string, netValue: number, revenue: number, expense: number}>, clients: Array<{name: string, netValue: number, revenue: number, expense: number}>, categories: Array<{name: string, netValue: number, revenue: number, expense: number}>}}
   */
  generateRankings(data) {
    if (!data || !Array.isArray(data) || data.length === 0) {
      return { projects: [], clients: [], categories: [] };
    }

    const projects = this._computeRanking(data, 'projeto', 10);
    const clients = this._computeRanking(data, 'cliente', 10);
    const categories = this._computeRanking(data, 'categoria', 10);

    return { projects, clients, categories };
  },

  /**
   * Compute ranking for a given dimension (field).
   * Groups records by the field value, sums revenue and expense,
   * calculates net value, and returns top N entries sorted by net value descending.
   * @param {FinancialRecord[]} data - Financial records
   * @param {string} field - Field name to group by
   * @param {number} limit - Maximum entries to return
   * @returns {Array<{name: string, netValue: number, revenue: number, expense: number}>}
   * @private
   */
  _computeRanking(data, field, limit) {
    const groups = {};

    for (const record of data) {
      const key = record[field] || '(sem nome)';
      if (!groups[key]) {
        groups[key] = { name: key, revenue: 0, expense: 0 };
      }

      if (record.tipo === 'receita') {
        groups[key].revenue += record.valor;
      } else if (record.tipo === 'despesa') {
        groups[key].expense += record.valor;
      }
    }

    // Calculate net value and sort descending
    const ranked = Object.values(groups).map(g => ({
      name: g.name,
      netValue: g.revenue - g.expense,
      revenue: g.revenue,
      expense: g.expense
    }));

    ranked.sort((a, b) => b.netValue - a.netValue);

    return ranked.slice(0, limit);
  },

  /**
   * Detect overdue payment alerts.
   * Finds records with status='vencido', groups by client/project,
   * and returns alert objects.
   * @param {FinancialRecord[]} data - Financial records
   * @returns {Array<{type: string, message: string, value: number, client: string, project: string}>}
   */
  detectOverdueAlerts(data) {
    if (!data || !Array.isArray(data) || data.length === 0) {
      return [];
    }

    const overdueRecords = data.filter(r => r.status === 'Pendente');
    if (overdueRecords.length === 0) return [];

    // Group by client + project combination
    const groups = {};
    for (const record of overdueRecords) {
      const client = record.cliente || '(sem cliente)';
      const project = record.projeto || '(sem projeto)';
      const key = `${client}|${project}`;

      if (!groups[key]) {
        groups[key] = { client, project, totalValue: 0, count: 0 };
      }
      groups[key].totalValue += record.valor;
      groups[key].count++;
    }

    // Generate alert objects
    const alerts = Object.values(groups).map(g => ({
      type: 'overdue',
      message: `${g.count} pagamento(s) vencido(s) - ${g.client} / ${g.project}: R$ ${g.totalValue.toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`,
      value: g.totalValue,
      client: g.client,
      project: g.project
    }));

    // Sort by value descending (highest overdue first)
    alerts.sort((a, b) => b.value - a.value);

    return alerts;
  },

  /**
   * Get top N clients by total revenue.
   * @param {FinancialRecord[]} data - Financial records
   * @param {number} limit - Number of top clients to return
   * @returns {Array<{name: string, netValue: number, revenue: number, expense: number}>}
   */
  getTopClients(data, limit) {
    if (!data || !Array.isArray(data) || data.length === 0) return [];

    const clientRevenue = {};
    for (const record of data) {
      if (record.tipo === 'receita') {
        const client = record.cliente || '(sem cliente)';
        if (!clientRevenue[client]) {
          clientRevenue[client] = { name: client, revenue: 0, expense: 0 };
        }
        clientRevenue[client].revenue += record.valor;
      }
    }

    // Also accumulate expenses for context
    for (const record of data) {
      if (record.tipo === 'despesa') {
        const client = record.cliente || '(sem cliente)';
        if (clientRevenue[client]) {
          clientRevenue[client].expense += record.valor;
        }
      }
    }

    const ranked = Object.values(clientRevenue).map(c => ({
      name: c.name,
      revenue: c.revenue,
      expense: c.expense,
      netValue: c.revenue - c.expense
    }));

    // Sort by revenue descending
    ranked.sort((a, b) => b.revenue - a.revenue);

    return ranked.slice(0, limit);
  },

  /**
   * Get top N categories by total expense.
   * @param {FinancialRecord[]} data - Financial records
   * @param {number} limit - Number of top categories to return
   * @returns {Array<{name: string, netValue: number, revenue: number, expense: number}>}
   */
  getTopExpenseCategories(data, limit) {
    if (!data || !Array.isArray(data) || data.length === 0) return [];

    const categoryExpense = {};
    for (const record of data) {
      if (record.tipo === 'despesa') {
        const category = record.categoria || '(sem categoria)';
        if (!categoryExpense[category]) {
          categoryExpense[category] = { name: category, expense: 0, revenue: 0 };
        }
        categoryExpense[category].expense += record.valor;
      }
    }

    // Also accumulate revenue for context
    for (const record of data) {
      if (record.tipo === 'receita') {
        const category = record.categoria || '(sem categoria)';
        if (categoryExpense[category]) {
          categoryExpense[category].revenue += record.valor;
        }
      }
    }

    const ranked = Object.values(categoryExpense).map(c => ({
      name: c.name,
      expense: c.expense,
      revenue: c.revenue,
      netValue: c.revenue - c.expense
    }));

    // Sort by expense descending (highest expense first)
    ranked.sort((a, b) => b.expense - a.expense);

    return ranked.slice(0, limit);
  },

  /**
   * Calculate 3-month moving direction trend.
   * Looks at the last 3 months of net totals (revenue - expense).
   * Returns '↑' if increasing, '↓' if decreasing, '→' if stable (change < 5%).
   * @param {Array<{month: string, total: number}>} monthlyTotals - Monthly net totals sorted chronologically
   * @returns {string} '↑'|'↓'|'→'
   */
  calculateTrendDirection(monthlyTotals) {
    if (!monthlyTotals || monthlyTotals.length < 3) {
      return '→'; // Not enough data for trend
    }

    // Take the last 3 months
    const last3 = monthlyTotals.slice(-3);
    const first = last3[0].total;
    const last = last3[2].total;

    // If first value is zero, determine direction by last value sign
    if (first === 0) {
      if (last > 0) return '↑';
      if (last < 0) return '↓';
      return '→';
    }

    const percentChange = ((last - first) / Math.abs(first)) * 100;

    if (percentChange >= 5) return '↑';
    if (percentChange <= -5) return '↓';
    return '→';
  },

  /**
   * Perform simple linear regression (least squares) on an array of {x, y} points.
   * Computes slope, intercept, and R² (coefficient of determination).
   * @param {Array<{x: number, y: number}>} points - Data points for regression
   * @returns {{slope: number, intercept: number, rSquared: number}}
   * @private
   */
  _linearRegression(points) {
    const n = points.length;
    if (n === 0) return { slope: 0, intercept: 0, rSquared: 0 };
    if (n === 1) return { slope: 0, intercept: points[0].y, rSquared: 0 };

    let sumX = 0, sumY = 0, sumXY = 0, sumX2 = 0;
    for (const p of points) {
      sumX += p.x;
      sumY += p.y;
      sumXY += p.x * p.y;
      sumX2 += p.x * p.x;
    }

    const meanX = sumX / n;
    const meanY = sumY / n;
    const denominator = sumX2 - (sumX * sumX) / n;

    let slope = 0;
    let intercept = meanY;

    if (denominator !== 0) {
      slope = (sumXY - (sumX * sumY) / n) / denominator;
      intercept = meanY - slope * meanX;
    }

    // Calculate R² (coefficient of determination)
    let ssRes = 0, ssTot = 0;
    for (const p of points) {
      const predicted = slope * p.x + intercept;
      ssRes += (p.y - predicted) ** 2;
      ssTot += (p.y - meanY) ** 2;
    }

    const rSquared = ssTot === 0 ? 0 : 1 - (ssRes / ssTot);

    return { slope, intercept, rSquared };
  },

  /**
   * Generate 3-month linear trend forecast.
   * Groups data by month, computes monthly net totals (revenue - expense),
   * applies linear regression, and projects next 3 months.
   * Requires at least 3 months of historical data.
   * @param {FinancialRecord[]} data - Financial records
   * @returns {{months: Array<{month: string, projected: number}>, slope: number, intercept: number, confidence: string}|null}
   *   Forecast result or null if insufficient data (fewer than 3 months)
   */
  generateForecast(data) {
    if (!data || !Array.isArray(data) || data.length === 0) return null;

    // Get monthly net totals
    const monthlyTotals = this._getMonthlyNetTotals(data);

    // Require at least 3 months of data
    if (monthlyTotals.length < 3) return null;

    // Build regression points: x = month index (0-based), y = net total
    const points = monthlyTotals.map((entry, index) => ({
      x: index,
      y: entry.total
    }));

    const { slope, intercept, rSquared } = this._linearRegression(points);

    // Determine confidence based on R²
    let confidence;
    if (rSquared > 0.8) {
      confidence = 'high';
    } else if (rSquared > 0.5) {
      confidence = 'medium';
    } else {
      confidence = 'low';
    }

    // Project next 3 months
    const lastMonthKey = monthlyTotals[monthlyTotals.length - 1].month;
    const [lastYear, lastMonth] = lastMonthKey.split('-').map(Number);

    const forecastMonths = [];
    for (let i = 1; i <= 3; i++) {
      const futureIndex = monthlyTotals.length - 1 + i;
      const projected = slope * futureIndex + intercept;

      // Calculate the future month key
      let futureMonth = lastMonth + i;
      let futureYear = lastYear;
      while (futureMonth > 12) {
        futureMonth -= 12;
        futureYear++;
      }
      const monthKey = `${futureYear}-${String(futureMonth).padStart(2, '0')}`;

      forecastMonths.push({ month: monthKey, projected });
    }

    return {
      months: forecastMonths,
      slope,
      intercept,
      confidence
    };
  },

  /**
   * Detect critical financial conditions from monthly changes.
   * Critical conditions: revenue drop >20% or expense increase >30% month-over-month.
   * @param {{expenseChanges: Array, revenueChanges: Array}} changes - Monthly changes from calculateMonthlyChanges
   * @returns {Array<{month: string, type: string, change: number, message: string, severity: string}>} Critical alerts
   */
  detectCriticalConditions(changes) {
    if (!changes) return [];

    const criticalAlerts = [];

    // Check revenue drops > 20%
    if (changes.revenueChanges && Array.isArray(changes.revenueChanges)) {
      for (const entry of changes.revenueChanges) {
        if (entry.change !== null && entry.change < -20) {
          criticalAlerts.push({
            month: entry.month,
            type: 'revenue_drop',
            change: entry.change,
            message: `Receita caiu ${Math.abs(entry.change).toFixed(1)}% em ${entry.month}`,
            severity: 'critical'
          });
        }
      }
    }

    // Check expense increases > 30%
    if (changes.expenseChanges && Array.isArray(changes.expenseChanges)) {
      for (const entry of changes.expenseChanges) {
        if (entry.change !== null && entry.change > 30) {
          criticalAlerts.push({
            month: entry.month,
            type: 'expense_increase',
            change: entry.change,
            message: `Despesas aumentaram ${entry.change.toFixed(1)}% em ${entry.month}`,
            severity: 'critical'
          });
        }
      }
    }

    return criticalAlerts;
  },

  /**
   * Compute monthly net totals (revenue - expense) for trend calculation.
   * @param {FinancialRecord[]} data - Financial records
   * @returns {Array<{month: string, total: number}>} Sorted monthly net totals
   * @private
   */
  _getMonthlyNetTotals(data) {
    const monthlyNet = {};

    for (const record of data) {
      if (!record.data || !(record.data instanceof Date) || isNaN(record.data.getTime())) continue;

      const year = record.data.getFullYear();
      const month = String(record.data.getMonth() + 1).padStart(2, '0');
      const monthKey = `${year}-${month}`;

      if (!monthlyNet[monthKey]) {
        monthlyNet[monthKey] = 0;
      }

      if (record.tipo === 'receita') {
        monthlyNet[monthKey] += record.valor;
      } else if (record.tipo === 'despesa') {
        monthlyNet[monthKey] -= record.valor;
      }
    }

    // Sort by month key and return as array
    return Object.entries(monthlyNet)
      .sort(([a], [b]) => a.localeCompare(b))
      .map(([month, total]) => ({ month, total }));
  }
};

// ============================================================
// SECTION: Export Module
// ============================================================

/** @type {Object} ExportModule - PDF and image export */
const ExportModule = {
  /**
   * Build PDF header with filter summary and export date.
   * Returns an array of header lines to be rendered in the PDF.
   * @param {FilterCriteria} filters - Active filters
   * @returns {string[]} Header lines for PDF
   */
  buildExportHeader(filters) {
    const lines = [];
    lines.push('Dashboard Financeiro - DOit ERP');

    // Export date in DD/MM/YYYY format
    const now = new Date();
    const day = String(now.getDate()).padStart(2, '0');
    const month = String(now.getMonth() + 1).padStart(2, '0');
    const year = now.getFullYear();
    lines.push(`Exportado em: ${day}/${month}/${year}`);

    // Active filters summary
    if (filters && Object.keys(filters).length > 0) {
      const filterLabels = {
        dataInicio: 'Período de',
        dataFim: 'Período até',
        projeto: 'Projeto',
        cliente: 'Cliente',
        categoria: 'Categoria',
        centroCusto: 'Centro de custo',
        status: 'Status',
        responsavel: 'Responsável'
      };

      const activeList = [];
      for (const [key, value] of Object.entries(filters)) {
        if (value != null && value !== '') {
          const label = filterLabels[key] || key;
          let displayValue = value;
          if (value instanceof Date) {
            const d = String(value.getDate()).padStart(2, '0');
            const m = String(value.getMonth() + 1).padStart(2, '0');
            const y = value.getFullYear();
            displayValue = `${d}/${m}/${y}`;
          }
          activeList.push(`${label}: ${displayValue}`);
        }
      }

      if (activeList.length > 0) {
        lines.push(`Filtros ativos: ${activeList.join(', ')}`);
      } else {
        lines.push('Sem filtros aplicados');
      }
    } else {
      lines.push('Sem filtros aplicados');
    }

    return lines;
  },

  /**
   * Trigger file download in the browser.
   * Creates a temporary <a> element, sets href to blob URL, triggers click, revokes URL.
   * @param {Blob} blob - File content
   * @param {string} filename - Download filename
   */
  triggerDownload(blob, filename) {
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    a.style.display = 'none';
    document.body.appendChild(a);
    a.click();
    // Clean up
    setTimeout(() => {
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
    }, 100);
  },

  /**
   * Generate PDF with all visible dashboard content.
   * Captures KPIs, charts (as canvas images), and table data.
   * Shows progress indicator during generation.
   * Displays error message with retry option on failure.
   * @returns {Promise<void>}
   */
  async generatePDF() {
    // Emit loading start
    EventBus.emit('loading:start', { stage: 'exporting' });

    try {
      // Access jsPDF constructor from global
      const { jsPDF } = window.jspdf;
      const doc = new jsPDF({ orientation: 'landscape', format: 'a4' });

      const pageWidth = doc.internal.pageSize.getWidth();
      const pageHeight = doc.internal.pageSize.getHeight();
      const margin = 15;
      let yPos = margin;

      // --- Header ---
      const filters = DataLayer.getActiveFilters();
      const headerLines = this.buildExportHeader(filters);

      doc.setFontSize(16);
      doc.setFont(undefined, 'bold');
      doc.text(headerLines[0], margin, yPos);
      yPos += 8;

      doc.setFontSize(10);
      doc.setFont(undefined, 'normal');
      for (let i = 1; i < headerLines.length; i++) {
        doc.text(headerLines[i], margin, yPos);
        yPos += 5;
      }
      yPos += 8;

      // --- KPI Values ---
      doc.setFontSize(12);
      doc.setFont(undefined, 'bold');
      doc.text('Indicadores (KPIs)', margin, yPos);
      yPos += 7;

      const data = DataLayer.getFilteredData(filters);
      const kpis = KPICalculator.calculateAll(data);

      const kpiLabels = [
        { key: 'totalRecebido', label: 'Total Recebido' },
        { key: 'totalPago', label: 'Total Pago' },
        { key: 'saldo', label: 'Saldo' },
        { key: 'recebimentosAberto', label: 'Recebimentos em Aberto' },
        { key: 'pagamentosAberto', label: 'Pagamentos em Aberto' },
        { key: 'fluxoMes', label: 'Fluxo do Mês' },
        { key: 'ticketMedio', label: 'Ticket Médio' },
        { key: 'quantidadeLancamentos', label: 'Qtd. Lançamentos' }
      ];

      doc.setFontSize(9);
      doc.setFont(undefined, 'normal');

      // Render KPIs in two columns
      const colWidth = (pageWidth - margin * 2) / 2;
      for (let i = 0; i < kpiLabels.length; i++) {
        const kpi = kpiLabels[i];
        const value = kpis[kpi.key];
        const formatted = kpi.key === 'quantidadeLancamentos'
          ? KPICalculator.formatInteger(value)
          : KPICalculator.formatBRL(value);

        const col = i % 2;
        const xPos = margin + col * colWidth;

        if (col === 0 && i > 0) {
          yPos += 5;
        }

        doc.text(`${kpi.label}: ${formatted}`, xPos, yPos);
      }
      yPos += 12;

      // --- Charts as images ---
      doc.setFontSize(12);
      doc.setFont(undefined, 'bold');
      doc.text('Gráficos', margin, yPos);
      yPos += 7;

      const chartIds = CHART_CONFIGS.map(c => c.id);
      const chartWidth = (pageWidth - margin * 2 - 10) / 2;
      const chartHeight = 55;
      let chartCol = 0;

      for (const chartId of chartIds) {
        const chartInstance = ChartEngine._instances[chartId];
        if (!chartInstance) continue;

        // Check if we need a new page
        if (yPos + chartHeight > pageHeight - margin) {
          doc.addPage();
          yPos = margin;
        }

        try {
          const imgData = chartInstance.toBase64Image();
          const xPos = margin + chartCol * (chartWidth + 10);
          doc.addImage(imgData, 'PNG', xPos, yPos, chartWidth, chartHeight);

          chartCol++;
          if (chartCol >= 2) {
            chartCol = 0;
            yPos += chartHeight + 5;
          }
        } catch (chartErr) {
          // Skip charts that fail to export
          console.warn(`Failed to export chart ${chartId}:`, chartErr);
        }
      }

      // If last row had only one chart, move yPos down
      if (chartCol !== 0) {
        yPos += chartHeight + 5;
      }

      // --- Table data ---
      // Add a new page for table data
      doc.addPage();
      yPos = margin;

      doc.setFontSize(12);
      doc.setFont(undefined, 'bold');
      doc.text('Dados Financeiros', margin, yPos);
      yPos += 7;

      doc.setFontSize(8);
      doc.setFont(undefined, 'normal');

      // Table header
      const tableHeaders = ['Projeto', 'Cliente', 'Categoria', 'Valor', 'Tipo', 'Data', 'Status'];
      const colWidths = [40, 35, 30, 30, 20, 25, 25];
      let xOffset = margin;

      doc.setFont(undefined, 'bold');
      for (let i = 0; i < tableHeaders.length; i++) {
        doc.text(tableHeaders[i], xOffset, yPos);
        xOffset += colWidths[i];
      }
      yPos += 5;
      doc.setFont(undefined, 'normal');

      // Table rows (limit to first 50 rows for PDF size)
      const tableData = data.slice(0, 50);
      for (const record of tableData) {
        if (yPos > pageHeight - margin) {
          doc.addPage();
          yPos = margin;

          // Re-draw header on new page
          xOffset = margin;
          doc.setFont(undefined, 'bold');
          for (let i = 0; i < tableHeaders.length; i++) {
            doc.text(tableHeaders[i], xOffset, yPos);
            xOffset += colWidths[i];
          }
          yPos += 5;
          doc.setFont(undefined, 'normal');
        }

        xOffset = margin;
        const rowValues = [
          (record.projeto || '').substring(0, 18),
          (record.cliente || '').substring(0, 16),
          (record.categoria || '').substring(0, 14),
          KPICalculator.formatBRL(record.valor),
          record.tipo || '',
          record.data ? TableModule.formatDate(record.data) : '',
          record.status || ''
        ];

        for (let i = 0; i < rowValues.length; i++) {
          doc.text(String(rowValues[i]), xOffset, yPos);
          xOffset += colWidths[i];
        }
        yPos += 4;
      }

      // If there are more records, add a note
      if (data.length > 50) {
        yPos += 3;
        doc.setFontSize(8);
        doc.setFont(undefined, 'italic');
        doc.text(`... e mais ${data.length - 50} registros (total: ${data.length})`, margin, yPos);
      }

      // --- Save PDF ---
      const pdfBlob = doc.output('blob');
      const now2 = new Date();
      const filename = `dashboard-financeiro-${now2.getFullYear()}${String(now2.getMonth() + 1).padStart(2, '0')}${String(now2.getDate()).padStart(2, '0')}.pdf`;

      // Emit loading end
      EventBus.emit('loading:end', {});

      // Trigger download
      this.triggerDownload(pdfBlob, filename);

    } catch (error) {
      // Emit loading end
      EventBus.emit('loading:end', {});

      // Emit error with retry action
      EventBus.emit('error:occurred', {
        type: 'export',
        stage: 'exporting',
        message: 'Erro ao gerar PDF. Tente novamente.',
        action: 'retry'
      });

      console.error('PDF export error:', error);
    }
  }
};

// Wire Export button
document.addEventListener('DOMContentLoaded', () => {
  const btnExportPdf = document.getElementById('btn-export-pdf');
  if (btnExportPdf) {
    btnExportPdf.addEventListener('click', () => ExportModule.generatePDF());
  }
});

// ============================================================
// SECTION: UI Controller & Initialization
// ============================================================

// ============================================================
// SECTION: Período Date Range Picker
// ============================================================

const PeriodoPicker = {
  _inputEl: null,
  _pickerEl: null,
  _onChange: null,
  _startDate: null,
  _endDate: null,
  _selecting: false, // true = selecting end date
  _viewYear: new Date().getFullYear(),
  _viewMonth: new Date().getMonth(),

  init(inputEl, pickerEl, onChange) {
    this._inputEl = inputEl;
    this._pickerEl = pickerEl;
    this._onChange = onChange;

    // Toggle picker on input click
    inputEl.addEventListener('click', (e) => {
      e.stopPropagation();
      if (pickerEl.hidden) {
        this._open();
      } else {
        this._close();
      }
    });

    // Navigation
    document.getElementById('periodo-prev-month').addEventListener('click', (e) => {
      e.stopPropagation();
      this._viewMonth--;
      if (this._viewMonth < 0) { this._viewMonth = 11; this._viewYear--; }
      this._renderDays();
    });
    document.getElementById('periodo-next-month').addEventListener('click', (e) => {
      e.stopPropagation();
      this._viewMonth++;
      if (this._viewMonth > 11) { this._viewMonth = 0; this._viewYear++; }
      this._renderDays();
    });

    // Click on month name to show month picker
    document.getElementById('periodo-month-name').addEventListener('click', (e) => {
      e.stopPropagation();
      this._renderMonthGrid();
    });

    // Click on year name to show year picker
    document.getElementById('periodo-year-name').addEventListener('click', (e) => {
      e.stopPropagation();
      this._toggleYearGrid();
    });

    // Clear button
    document.getElementById('periodo-clear').addEventListener('click', (e) => {
      e.stopPropagation();
      this.clear();
      if (this._onChange) this._onChange();
      this._close();
    });

    // Apply button
    document.getElementById('periodo-apply').addEventListener('click', (e) => {
      e.stopPropagation();
      this._updateInput();
      if (this._onChange) this._onChange();
      this._close();
    });

    // Close on outside click
    document.addEventListener('click', (e) => {
      if (!pickerEl.hidden && !pickerEl.contains(e.target) && e.target !== inputEl) {
        this._close();
      }
    });

    // Prevent picker clicks from closing
    pickerEl.addEventListener('click', (e) => e.stopPropagation());
  },

  _open() {
    this._pickerEl.hidden = false;
    if (this._startDate) {
      this._viewYear = this._startDate.getFullYear();
      this._viewMonth = this._startDate.getMonth();
    }
    this._renderDays();
  },

  _close() {
    this._pickerEl.hidden = true;
  },

  _renderDays() {
    const monthNameEl = document.getElementById('periodo-month-name');
    const yearNameEl = document.getElementById('periodo-year-name');
    const container = document.getElementById('periodo-days');
    const weekdays = document.getElementById('periodo-weekdays');
    const yearGrid = document.getElementById('periodo-year-grid');
    const monthGrid = document.getElementById('periodo-month-grid');
    const monthNames = ['Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho', 'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro'];

    if (monthNameEl) monthNameEl.textContent = monthNames[this._viewMonth];
    if (yearNameEl) yearNameEl.textContent = this._viewYear;
    if (yearGrid) yearGrid.hidden = true;
    if (monthGrid) monthGrid.hidden = true;
    if (weekdays) weekdays.hidden = false;
    if (container) container.hidden = false;

    const firstDay = new Date(this._viewYear, this._viewMonth, 1).getDay();
    const daysInMonth = new Date(this._viewYear, this._viewMonth + 1, 0).getDate();
    const today = new Date();
    today.setHours(0, 0, 0, 0);

    let html = '';
    // Empty cells for days before 1st
    for (let i = 0; i < firstDay; i++) {
      html += '<button class="periodo-day empty" disabled></button>';
    }

    for (let d = 1; d <= daysInMonth; d++) {
      const date = new Date(this._viewYear, this._viewMonth, d);
      date.setHours(0, 0, 0, 0);
      let cls = 'periodo-day';

      if (date.getTime() === today.getTime()) cls += ' today';

      if (this._startDate && date.getTime() === this._startDate.getTime()) cls += ' selected';
      if (this._endDate && date.getTime() === this._endDate.getTime()) cls += ' selected';
      if (this._startDate && this._endDate &&
          date > this._startDate && date < this._endDate) cls += ' in-range';

      html += `<button class="${cls}" data-date="${this._viewYear}-${String(this._viewMonth + 1).padStart(2, '0')}-${String(d).padStart(2, '0')}">${d}</button>`;
    }

    container.innerHTML = html;

    // Day click handlers
    container.querySelectorAll('.periodo-day:not(.empty)').forEach(btn => {
      btn.addEventListener('click', () => {
        const dateStr = btn.dataset.date;
        const clicked = new Date(dateStr + 'T00:00:00');

        if (!this._startDate || this._selecting === false) {
          // First click: set start
          this._startDate = clicked;
          this._endDate = null;
          this._selecting = true;
        } else {
          // Second click: set end
          if (clicked < this._startDate) {
            this._endDate = this._startDate;
            this._startDate = clicked;
          } else {
            this._endDate = clicked;
          }
          this._selecting = false;
        }
        this._renderDays();
      });
    });
  },

  _updateInput() {
    if (!this._inputEl) return;
    if (!this._startDate) {
      this._inputEl.value = '';
      return;
    }
    const fmt = (d) => `${String(d.getDate()).padStart(2, '0')}/${String(d.getMonth() + 1).padStart(2, '0')}/${d.getFullYear()}`;
    if (this._endDate && this._startDate.getTime() !== this._endDate.getTime()) {
      this._inputEl.value = `${fmt(this._startDate)} - ${fmt(this._endDate)}`;
    } else {
      this._inputEl.value = fmt(this._startDate);
    }
  },

  getRange() {
    const result = { start: null, end: null };
    if (this._startDate) {
      result.start = new Date(this._startDate);
      result.start.setHours(0, 0, 0, 0);
    }
    if (this._endDate) {
      result.end = new Date(this._endDate);
      result.end.setHours(23, 59, 59, 999);
    } else if (this._startDate) {
      // Single day selected
      result.end = new Date(this._startDate);
      result.end.setHours(23, 59, 59, 999);
    }
    return result;
  },

  _toggleYearGrid() {
    const grid = document.getElementById('periodo-year-grid');
    const monthGrid = document.getElementById('periodo-month-grid');
    const daysContainer = document.getElementById('periodo-days');
    const weekdays = document.getElementById('periodo-weekdays');

    if (!grid) return;

    if (grid.hidden) {
      // Show year grid
      grid.hidden = false;
      if (monthGrid) monthGrid.hidden = true;
      if (daysContainer) daysContainer.hidden = true;
      if (weekdays) weekdays.hidden = true;
      this._renderYearGrid();
    } else {
      // Back to days
      grid.hidden = true;
      this._renderDays();
    }
  },

  _renderYearGrid() {
    const grid = document.getElementById('periodo-year-grid');
    if (!grid) return;

    const currentYear = new Date().getFullYear();
    let html = '';
    for (let y = currentYear - 5; y <= currentYear + 5; y++) {
      const cls = y === this._viewYear ? 'periodo-year-btn active' : 'periodo-year-btn';
      html += `<button class="${cls}" data-year="${y}">${y}</button>`;
    }
    grid.innerHTML = html;

    grid.querySelectorAll('.periodo-year-btn').forEach(btn => {
      btn.addEventListener('click', (e) => {
        e.stopPropagation();
        this._viewYear = parseInt(btn.dataset.year);
        this._renderDays();
      });
    });
  },

  _renderMonthGrid() {
    const grid = document.getElementById('periodo-year-grid');
    const monthGrid = document.getElementById('periodo-month-grid');
    const daysContainer = document.getElementById('periodo-days');
    const weekdays = document.getElementById('periodo-weekdays');

    if (grid) grid.hidden = true;
    if (monthGrid) monthGrid.hidden = false;
    if (daysContainer) daysContainer.hidden = true;
    if (weekdays) weekdays.hidden = true;

    const monthNames = ['Jan','Fev','Mar','Abr','Mai','Jun','Jul','Ago','Set','Out','Nov','Dez'];
    let html = '';
    for (let m = 0; m < 12; m++) {
      const cls = m === this._viewMonth ? 'periodo-year-btn active' : 'periodo-year-btn';
      html += `<button class="${cls}" data-month="${m}">${monthNames[m]}</button>`;
    }
    if (monthGrid) monthGrid.innerHTML = html;

    monthGrid.querySelectorAll('.periodo-year-btn').forEach(btn => {
      btn.addEventListener('click', (e) => {
        e.stopPropagation();
        this._viewMonth = parseInt(btn.dataset.month);
        this._renderDays();
      });
    });
  },

  clear() {
    this._startDate = null;
    this._endDate = null;
    this._selecting = false;
    if (this._inputEl) this._inputEl.value = '';
  }
};

/** @type {Object} UIController - DOM interactions and app lifecycle */
const UIController = {
  /**
   * Filter element ID to DataLayer column name mapping.
   * @type {Object<string, string>}
   */
  _filterMapping: {
    'filter-projeto': 'projeto',
    'filter-cliente': 'cliente',
    'filter-tipo': 'tipoPagamento',
    'filter-nivel1': 'tipoClassif',
    'filter-nivel2': 'nivel1',
    'filter-nivel3': 'nivel2',
    'filter-centro-custo': 'centroCusto',
    'filter-conta': 'conta',
    'filter-status': 'status'
  },

  /**
   * Set tracking which sections have been rendered via lazy rendering.
   * Sections are added here once their content is rendered.
   * @type {Set<string>}
   */
  _sectionRendered: new Set(),

  /**
   * IntersectionObserver instance for lazy rendering.
   * @type {IntersectionObserver|null}
   * @private
   */
  _lazyObserver: null,

  /**
   * Cached data reference for deferred section rendering.
   * @type {FinancialRecord[]|null}
   * @private
   */
  _currentData: null,

  /**
   * Sections that should be rendered immediately on data load (priority sections).
   * KPI cards (section-dashboard) and the first chart section (section-financeiro).
   * @type {string[]}
   * @private
   */
  _prioritySections: ['section-dashboard', 'section-financeiro'],

  /**
   * Sections that are deferred until scrolled into view.
   * @type {string[]}
   * @private
   */
  _deferredSections: ['section-projetos', 'section-projeto-detalhe', 'section-relatorios', 'section-insights'],

  /**
   * Initialize lazy rendering using IntersectionObserver.
   * Observes each .dashboard-section element with a 200px rootMargin.
   * When a deferred section enters the threshold, its content is rendered
   * and the observer stops watching it.
   *
   * Validates: Requirements 12.2, 12.4
   */
  initLazyRendering() {
    // Clean up any existing observer
    if (this._lazyObserver) {
      this._lazyObserver.disconnect();
      this._lazyObserver = null;
    }

    // If IntersectionObserver is not supported, render all sections immediately
    if (typeof IntersectionObserver === 'undefined') {
      this._deferredSections.forEach(sectionId => {
        this.renderSection(sectionId);
      });
      return;
    }

    this._lazyObserver = new IntersectionObserver((entries) => {
      entries.forEach(entry => {
        if (entry.isIntersecting) {
          const sectionId = entry.target.id;
          if (!this._sectionRendered.has(sectionId)) {
            this.renderSection(sectionId);
          }
          // Unobserve after rendering
          this._lazyObserver.unobserve(entry.target);
        }
      });
    }, {
      rootMargin: `${PERFORMANCE.LAZY_RENDER_OFFSET}px`,
      threshold: 0
    });

    // Observe only deferred sections
    this._deferredSections.forEach(sectionId => {
      const sectionEl = document.getElementById(sectionId);
      if (sectionEl && !this._sectionRendered.has(sectionId)) {
        this._lazyObserver.observe(sectionEl);
      }
    });
  },

  /**
   * Render content for a specific section by its ID.
   * For chart sections: initializes the charts in that section.
   * For table sections: renders the tables.
   * For insights/relatorios: renders insights.
   *
   * @param {string} sectionId - The section element ID to render
   */
  renderSection(sectionId) {
    if (this._sectionRendered.has(sectionId)) return;

    const data = this._currentData || DataLayer.getData();
    const filteredData = DataLayer.getFilteredData(DataLayer.getActiveFilters());
    const renderData = Object.keys(DataLayer.getActiveFilters()).length > 0 ? filteredData : data;

    switch (sectionId) {
      case 'section-financeiro':
        // Initialize charts for the financeiro section
        ChartEngine.initCharts(renderData);
        break;

      case 'section-projetos':
        // All tables are now in this single section
        this.initTables(renderData);
        break;

      case 'section-projeto-detalhe':
        // Populate project selector
        ProjetoDetalhe.populateProjects();
        break;

      case 'section-relatorios':
        // Render report
        ReportModule.generateReport(renderData);
        break;

      case 'section-insights':
        // Render insights
        this.renderInsights(renderData);
        break;

      default:
        break;
    }

    this._sectionRendered.add(sectionId);
  },

  /**
   * Initialize the filter bar by attaching event listeners to all filter
   * inputs, selects, and the clear button.
   * Called once during app initialization.
   */
  initFilterBar() {
    // Date range picker (periodo)
    const periodoInput = document.getElementById('filter-periodo');
    const periodoPicker = document.getElementById('periodo-picker');

    if (periodoInput && periodoPicker) {
      // Initialize periodo picker module
      PeriodoPicker.init(periodoInput, periodoPicker, () => this.onFilterChange());
    }

    // Dropdown select filters
    for (const elementId of Object.keys(this._filterMapping)) {
      const selectEl = document.getElementById(elementId);
      if (selectEl) {
        selectEl.addEventListener('change', () => this.onFilterChange());
      }
    }

    // Cascading category filters (nivel1 → nivel2 → nivel3)
    const nivel1El = document.getElementById('filter-nivel1');
    const nivel2El = document.getElementById('filter-nivel2');
    const nivel3El = document.getElementById('filter-nivel3');
    const nivel2Group = document.getElementById('filter-group-nivel2');
    const nivel3Group = document.getElementById('filter-group-nivel3');

    if (nivel1El) {
      nivel1El.addEventListener('change', () => {
        const val = nivel1El.value;
        if (val) {
          // Show nivel2, populate with options within selected nivel1
          if (nivel2Group) nivel2Group.hidden = false;
          this._populateCascadeFilter('filter-nivel2', 'nivel1', r => r.tipoClassif === val);
        } else {
          // Hide nivel2 and nivel3
          if (nivel2Group) nivel2Group.hidden = true;
          if (nivel3Group) nivel3Group.hidden = true;
          if (nivel2El) nivel2El.value = '';
          if (nivel3El) nivel3El.value = '';
        }
      });
    }

    if (nivel2El) {
      nivel2El.addEventListener('change', () => {
        const val = nivel2El.value;
        if (val) {
          // Show nivel3, populate with options within selected nivel2
          if (nivel3Group) nivel3Group.hidden = false;
          const nivel1Val = nivel1El ? nivel1El.value : '';
          this._populateCascadeFilter('filter-nivel3', 'nivel2', r => r.nivel1 === val && (!nivel1Val || r.tipoClassif === nivel1Val));
        } else {
          // Hide nivel3
          if (nivel3Group) nivel3Group.hidden = true;
          if (nivel3El) nivel3El.value = '';
        }
      });
    }

    // Clear filters button
    const btnClear = document.getElementById('btn-clear-filters');
    if (btnClear) {
      btnClear.addEventListener('click', () => this.clearFilters());
    }
  },

  /**
   * Populate all dropdown filters with unique values from the loaded data.
   * Each select gets an option for every unique non-null value in its
   * corresponding column, plus the default "Todos" option.
   * Called when 'data:loaded' event fires.
   */
  populateFilters() {
    for (const [elementId, columnName] of Object.entries(this._filterMapping)) {
      const selectEl = document.getElementById(elementId);
      if (!selectEl) continue;

      // Get unique sorted values from DataLayer
      const values = DataLayer.getColumnValues(columnName);

      // Clear existing options (keep the default "Todos")
      selectEl.innerHTML = '<option value="">Todos</option>';

      // Add an option for each unique value
      for (const value of values) {
        const option = document.createElement('option');
        option.value = value;
        option.textContent = value;
        selectEl.appendChild(option);
      }
    }
  },

  /**
   * Populate a cascading filter dropdown with values from records matching a predicate.
   * @param {string} selectId - ID of the select element
   * @param {string} columnName - Column to extract unique values from
   * @param {function} predicate - Filter predicate to limit source records
   */
  _populateCascadeFilter(selectId, columnName, predicate) {
    const selectEl = document.getElementById(selectId);
    if (!selectEl) return;

    const allData = DataLayer.getData();
    const filtered = allData.filter(predicate);
    const values = [...new Set(filtered.map(r => r[columnName]).filter(Boolean))].sort((a, b) => a.localeCompare(b, 'pt-BR'));

    selectEl.innerHTML = '<option value="">Todos</option>';
    for (const value of values) {
      const option = document.createElement('option');
      option.value = value;
      option.textContent = value;
      selectEl.appendChild(option);
    }
  },

  /**
   * Read all current filter values from the DOM, build a FilterCriteria
   * object, and pass it to FilterEngine.applyFiltersDebounced.
   * Called whenever any filter input/select changes.
   */
  onFilterChange() {
    const criteria = {};

    // Date range from PeriodoPicker
    const range = PeriodoPicker.getRange();
    if (range.start) {
      criteria.dataInicio = range.start;
    }
    if (range.end) {
      criteria.dataFim = range.end;
    }

    // Dropdown filters
    for (const [elementId, columnName] of Object.entries(this._filterMapping)) {
      const selectEl = document.getElementById(elementId);
      if (selectEl && selectEl.value) {
        criteria[columnName] = selectEl.value;
      }
    }

    FilterEngine.applyFiltersDebounced(criteria);
  },

  /**
   * Reset all filter controls to their default state and call
   * DataLayer.clearFilters to restore the full dataset.
   */
  clearFilters() {
    // Reset periodo picker
    PeriodoPicker.clear();

    // Reset all select dropdowns to default (empty value = "Todos")
    for (const elementId of Object.keys(this._filterMapping)) {
      const selectEl = document.getElementById(elementId);
      if (selectEl) selectEl.value = '';
    }

    // Hide cascading category groups
    const nivel2Group = document.getElementById('filter-group-nivel2');
    const nivel3Group = document.getElementById('filter-group-nivel3');
    if (nivel2Group) nivel2Group.hidden = true;
    if (nivel3Group) nivel3Group.hidden = true;

    // Clear filters in DataLayer (emits 'filters:cleared')
    DataLayer.clearFilters();

    // Hide no-results message
    const noResults = document.getElementById('filter-no-results');
    if (noResults) noResults.hidden = true;
  },

  /**
   * Show or hide the "no results" message based on filtered data length.
   * @param {FinancialRecord[]} filteredData - The current filtered dataset
   */
  checkNoResults(filteredData) {
    const noResults = document.getElementById('filter-no-results');
    if (!noResults) return;

    if (!filteredData || filteredData.length === 0) {
      noResults.hidden = false;
    } else {
      noResults.hidden = true;
    }
  },

  // ---- Table UI Rendering ----

  /**
   * Internal state for each table: current page, sort column/direction, search query, and data.
   * @type {Object<string, {page: number, sortColumn: string|null, sortDirection: 'asc'|'desc', searchQuery: string, data: Array, filteredData: Array, columns: Array}>}
   */
  _tableState: {},

  /**
   * Track whether table event handlers have been set up (to avoid duplicate listeners).
   * @type {Object<string, boolean>}
   * @private
   */
  _tableHandlersInitialized: {},

  /**
   * Table configurations: defines columns and data extraction logic for each table.
   * @type {Object<string, {title: string, columns: Array<{key: string, label: string}>}>}
   */
  _tableConfigs: {
    'ultimos-lancamentos': {
      title: 'Últimos Lançamentos',
      columns: [
        { key: 'data', label: 'Data' },
        { key: 'descricao', label: 'Descrição' },
        { key: 'projeto', label: 'Projeto' },
        { key: 'cliente', label: 'De/Para' },
        { key: 'valor', label: 'Valor' },
        { key: 'tipoPagamento', label: 'Tipo' },
        { key: 'tipoClassif', label: '1ª Cat.' },
        { key: 'nivel1', label: '2ª Cat.' },
        { key: 'nivel2', label: '3ª Cat.' },
        { key: 'status', label: 'Status' }
      ]
    },
    'maiores-despesas': {
      title: 'Maiores Despesas',
      columns: [
        { key: 'data', label: 'Data' },
        { key: 'projeto', label: 'Projeto' },
        { key: 'cliente', label: 'De/Para' },
        { key: 'valor', label: 'Valor' },
        { key: 'tipoClassif', label: '1ª Cat.' },
        { key: 'nivel1', label: '2ª Cat.' },
        { key: 'nivel2', label: '3ª Cat.' }
      ]
    },
    'maiores-receitas': {
      title: 'Maiores Receitas',
      columns: [
        { key: 'data', label: 'Data' },
        { key: 'projeto', label: 'Projeto' },
        { key: 'cliente', label: 'De/Para' },
        { key: 'valor', label: 'Valor' },
        { key: 'tipoClassif', label: '1ª Cat.' },
        { key: 'nivel1', label: '2ª Cat.' },
        { key: 'nivel2', label: '3ª Cat.' }
      ]
    },
    'projetos-lucrativos': {
      title: 'Projetos Mais Lucrativos',
      columns: [
        { key: 'name', label: 'Projeto' },
        { key: 'revenue', label: 'Receita' },
        { key: 'expense', label: 'Despesa' },
        { key: 'netValue', label: 'Saldo' }
      ]
    },
    'clientes-faturamento': {
      title: 'Clientes com Maior Faturamento',
      columns: [
        { key: 'name', label: 'Cliente' },
        { key: 'revenue', label: 'Receita' },
        { key: 'expense', label: 'Despesa' },
        { key: 'netValue', label: 'Saldo' }
      ]
    },
    'todos-lancamentos': {
      title: 'Todos os Lançamentos',
      columns: [
        { key: 'data', label: 'Data' },
        { key: 'descricao', label: 'Descrição' },
        { key: 'projeto', label: 'Projeto' },
        { key: 'cliente', label: 'De/Para' },
        { key: 'valor', label: 'Valor' },
        { key: 'tipoPagamento', label: 'Tipo' },
        { key: 'tipoClassif', label: '1ª Cat.' },
        { key: 'nivel1', label: '2ª Cat.' },
        { key: 'nivel2', label: '3ª Cat.' },
        { key: 'status', label: 'Status' }
      ]
    }
  },

  /**
   * Initialize all 5 tables with their respective data subsets.
   * Sets up sort handlers, search handlers, and renders initial data.
   * @param {FinancialRecord[]} data - Full financial dataset (possibly filtered)
   */
  initTables(data) {
    const tableIds = Object.keys(this._tableConfigs);

    for (const tableId of tableIds) {
      const tableData = this.getTableData(tableId, data);
      const config = this._tableConfigs[tableId];

      // Initialize state for this table
      this._tableState[tableId] = {
        page: 1,
        sortColumn: 'data',
        sortDirection: tableId === 'ultimos-lancamentos' ? 'desc' : 'asc',
        searchQuery: '',
        data: tableData,
        filteredData: tableData,
        columns: config.columns
      };

      // Render the table
      this.renderTable(tableId);

      // Setup sort and search handlers (only once)
      if (!this._tableHandlersInitialized[tableId]) {
        this.setupTableSort(tableId);
        this.setupTableSearch(tableId);
        this._tableHandlersInitialized[tableId] = true;
      }
    }
  },

  /**
   * Extract the appropriate data subset for each table from the full dataset.
   * - 'ultimos-lancamentos': sort by date desc, take first 50
   * - 'maiores-despesas': filter tipo='despesa', sort by valor desc, take top 10
   * - 'maiores-receitas': filter tipo='receita', sort by valor desc, take top 10
   * - 'projetos-lucrativos': aggregate by projeto (revenue-expense), sort desc, top 10
   * - 'clientes-faturamento': aggregate by cliente (total revenue), sort desc, top 10
   *
   * @param {string} tableId - Table identifier
   * @param {FinancialRecord[]} allData - Full dataset
   * @returns {Array} Data subset for the table
   */
  getTableData(tableId, allData) {
    if (!allData || !Array.isArray(allData) || allData.length === 0) return [];

    switch (tableId) {
      case 'ultimos-lancamentos': {
        // Sort by date descending, take first 50
        const sorted = [...allData].sort((a, b) => {
          const dateA = a.data instanceof Date ? a.data.getTime() : new Date(a.data).getTime() || 0;
          const dateB = b.data instanceof Date ? b.data.getTime() : new Date(b.data).getTime() || 0;
          return dateB - dateA;
        });
        return sorted.slice(0, 50);
      }

      case 'maiores-despesas': {
        // Filter tipo='despesa', sort by valor desc, take top 10
        const despesas = allData.filter(r => r.tipo === 'despesa');
        const sorted = [...despesas].sort((a, b) => b.valor - a.valor);
        return sorted.slice(0, 10);
      }

      case 'maiores-receitas': {
        // Filter tipo='receita', sort by valor desc, take top 10
        const receitas = allData.filter(r => r.tipo === 'receita');
        const sorted = [...receitas].sort((a, b) => b.valor - a.valor);
        return sorted.slice(0, 10);
      }

      case 'projetos-lucrativos': {
        // Aggregate by projeto: revenue - expense, sort desc, top 10
        const groups = {};
        for (const record of allData) {
          const key = record.projeto || '(sem projeto)';
          if (!groups[key]) {
            groups[key] = { name: key, revenue: 0, expense: 0 };
          }
          if (record.tipo === 'receita') {
            groups[key].revenue += record.valor;
          } else if (record.tipo === 'despesa') {
            groups[key].expense += record.valor;
          }
        }
        const ranked = Object.values(groups).map(g => ({
          name: g.name,
          revenue: g.revenue,
          expense: g.expense,
          netValue: g.revenue - g.expense
        }));
        ranked.sort((a, b) => b.netValue - a.netValue);
        return ranked.slice(0, 10);
      }

      case 'clientes-faturamento': {
        // Aggregate by cliente: total revenue, sort desc, top 10
        const groups = {};
        for (const record of allData) {
          const key = record.cliente || '(sem cliente)';
          if (!groups[key]) {
            groups[key] = { name: key, revenue: 0, expense: 0 };
          }
          if (record.tipo === 'receita') {
            groups[key].revenue += record.valor;
          } else if (record.tipo === 'despesa') {
            groups[key].expense += record.valor;
          }
        }
        const ranked = Object.values(groups).map(g => ({
          name: g.name,
          revenue: g.revenue,
          expense: g.expense,
          netValue: g.revenue - g.expense
        }));
        // Sort by revenue (faturamento) descending
        ranked.sort((a, b) => b.revenue - a.revenue);
        return ranked.slice(0, 10);
      }

      default:
        // 'todos-lancamentos': return all records sorted by date desc
        if (tableId === 'todos-lancamentos') {
          return [...allData].sort((a, b) => {
            const da = a.data instanceof Date ? a.data : new Date(a.data);
            const db = b.data instanceof Date ? b.data : new Date(b.data);
            return db - da;
          });
        }
        return [];
    }
  },

  /**
   * Render a table with headers, rows, and pagination.
   * Uses the current state (page, sort, search) for the given tableId.
   * @param {string} tableId - Table identifier
   */
  renderTable(tableId) {
    const state = this._tableState[tableId];
    if (!state) return;

    const tableEl = document.getElementById(`table-${tableId}`);
    if (!tableEl) return;

    const config = this._tableConfigs[tableId];
    const columns = config.columns;

    // Get display data (already filtered by search)
    let displayData = state.filteredData;

    // Apply sort if active
    if (state.sortColumn) {
      displayData = this._sortTableData(displayData, state.sortColumn, state.sortDirection);
    }

    // Paginate
    const { rows, currentPage, totalPages } = TableModule.paginate(displayData, state.page);
    state.page = currentPage;

    // Render table header
    const thead = tableEl.querySelector('thead tr');
    if (thead) {
      thead.innerHTML = '';
      for (const col of columns) {
        const th = document.createElement('th');
        th.setAttribute('data-column', col.key);
        th.setAttribute('role', 'columnheader');
        th.setAttribute('aria-sort', 'none');
        th.style.cursor = 'pointer';
        th.setAttribute('tabindex', '0');

        // Column label
        const labelSpan = document.createElement('span');
        labelSpan.textContent = col.label;
        th.appendChild(labelSpan);

        // Sort direction indicator
        const sortIndicator = document.createElement('span');
        sortIndicator.className = 'sort-indicator';
        sortIndicator.setAttribute('aria-hidden', 'true');
        if (state.sortColumn === col.key) {
          sortIndicator.textContent = state.sortDirection === 'asc' ? ' ▲' : ' ▼';
          th.setAttribute('aria-sort', state.sortDirection === 'asc' ? 'ascending' : 'descending');
        } else {
          sortIndicator.textContent = '';
        }
        th.appendChild(sortIndicator);

        thead.appendChild(th);
      }
    }

    // Render table body
    const tbody = tableEl.querySelector('tbody');
    if (tbody) {
      tbody.innerHTML = '';

      for (const row of rows) {
        const tr = document.createElement('tr');
        // Add record ID for click-to-detail
        if (row.id) tr.dataset.recordId = row.id;
        // For aggregate tables, add name for drill-down popup
        if (!row.id && row.name) {
          tr.dataset.aggregateName = row.name;
          tr.dataset.aggregateTable = tableId;
          tr.style.cursor = 'pointer';
        }

        for (const col of columns) {
          const td = document.createElement('td');
          td.setAttribute('data-column', col.key);
          td.textContent = this._formatCellValue(col.key, row[col.key]);
          tr.appendChild(td);
        }

        tbody.appendChild(tr);
      }
    }

    // Update pagination controls
    this.updateTablePagination(tableId, currentPage, totalPages);

    // Show/hide empty message
    const tableContainer = tableEl.closest('.table-section') || tableEl.closest('.table-card');
    if (tableContainer) {
      const emptyMsg = tableContainer.querySelector('.table-empty');
      if (emptyMsg) {
        emptyMsg.hidden = displayData.length > 0;
      }
    }
  },

  /**
   * Sort table data by column, handling different data types.
   * @param {Array} data - Data to sort
   * @param {string} column - Column key
   * @param {'asc'|'desc'} direction - Sort direction
   * @returns {Array} Sorted data (new array)
   * @private
   */
  _sortTableData(data, column, direction) {
    const dir = direction === 'desc' ? -1 : 1;

    return [...data].sort((a, b) => {
      const valA = a[column];
      const valB = b[column];

      // Handle null/undefined
      if (valA == null && valB == null) return 0;
      if (valA == null) return 1;
      if (valB == null) return -1;

      // Numeric columns
      if (column === 'valor' || column === 'revenue' || column === 'expense' || column === 'netValue') {
        return (Number(valA) - Number(valB)) * dir;
      }

      // Date column
      if (column === 'data') {
        const dateA = valA instanceof Date ? valA.getTime() : new Date(valA).getTime();
        const dateB = valB instanceof Date ? valB.getTime() : new Date(valB).getTime();
        return (dateA - dateB) * dir;
      }

      // String comparison with locale-aware sorting
      return String(valA).localeCompare(String(valB), 'pt-BR') * dir;
    });
  },

  /**
   * Format a cell value for display based on column type.
   * Currency columns use BRL format, date columns use DD/MM/YYYY.
   * @param {string} columnKey - Column key
   * @param {*} value - Raw value
   * @returns {string} Formatted display value
   * @private
   */
  _formatCellValue(columnKey, value) {
    if (value == null || value === '') return '';

    // Currency columns
    if (columnKey === 'valor' || columnKey === 'revenue' || columnKey === 'expense' || columnKey === 'netValue') {
      return KPICalculator.formatBRL(Number(value));
    }

    // Date column
    if (columnKey === 'data') {
      return TableModule.formatDate(value);
    }

    return String(value);
  },

  /**
   * Add click handlers to column headers for sorting.
   * Clicking a header sorts ascending on first click, toggles to descending on subsequent clicks.
   * @param {string} tableId - Table identifier
   */
  setupTableSort(tableId) {
    const tableEl = document.getElementById(`table-${tableId}`);
    if (!tableEl) return;

    const thead = tableEl.querySelector('thead');
    if (!thead) return;

    // Use event delegation on thead
    thead.addEventListener('click', (e) => {
      const th = e.target.closest('th');
      if (!th) return;

      const column = th.getAttribute('data-column');
      if (!column) return;

      const state = this._tableState[tableId];
      if (!state) return;

      // Toggle sort direction
      if (state.sortColumn === column) {
        state.sortDirection = state.sortDirection === 'asc' ? 'desc' : 'asc';
      } else {
        state.sortColumn = column;
        state.sortDirection = 'asc';
      }

      // Reset to page 1 on sort change
      state.page = 1;

      // Re-render
      this.renderTable(tableId);
    });

    // Keyboard support for sort headers
    thead.addEventListener('keydown', (e) => {
      if (e.key === 'Enter' || e.key === ' ') {
        const th = e.target.closest('th');
        if (th) {
          e.preventDefault();
          th.click();
        }
      }
    });
  },

  /**
   * Add input handler with debounced search to the table's search field.
   * Performs case-insensitive partial match against all visible columns.
   * @param {string} tableId - Table identifier
   */
  setupTableSearch(tableId) {
    const searchInput = document.getElementById(`search-${tableId}`);
    if (!searchInput) return;

    let debounceTimer = null;

    searchInput.addEventListener('input', () => {
      const query = searchInput.value;

      if (debounceTimer) {
        clearTimeout(debounceTimer);
      }

      debounceTimer = setTimeout(() => {
        const state = this._tableState[tableId];
        if (!state) return;

        state.searchQuery = query;

        if (!query || query.trim() === '') {
          state.filteredData = state.data;
        } else {
          const lowerQuery = query.toLowerCase();
          const columns = state.columns;

          state.filteredData = state.data.filter(row => {
            return columns.some(col => {
              const value = row[col.key];
              if (value == null) return false;

              // Format for comparison based on type
              let displayValue;
              if (col.key === 'data' && value instanceof Date) {
                displayValue = TableModule.formatDate(value);
              } else if (col.key === 'valor' || col.key === 'revenue' || col.key === 'expense' || col.key === 'netValue') {
                displayValue = KPICalculator.formatBRL(Number(value));
              } else {
                displayValue = String(value);
              }

              return displayValue.toLowerCase().includes(lowerQuery);
            });
          });
        }

        // Reset to page 1 on search
        state.page = 1;

        // Re-render
        this.renderTable(tableId);
      }, PERFORMANCE.SEARCH_DEBOUNCE);
    });
  },

  /**
   * Update pagination controls for a table.
   * Shows current page, total pages, and enables/disables prev/next buttons.
   * @param {string} tableId - Table identifier
   * @param {number} currentPage - Current page number (1-based)
   * @param {number} totalPages - Total number of pages
   */
  updateTablePagination(tableId, currentPage, totalPages) {
    const tableEl = document.getElementById(`table-${tableId}`);
    if (!tableEl) return;

    const tableSection = tableEl.closest('.table-section') || tableEl.closest('.table-card');
    if (!tableSection) return;

    const paginationEl = tableSection.querySelector('.table-pagination');
    if (!paginationEl) return;

    const pageInfo = paginationEl.querySelector('.page-info');
    const btnPrev = paginationEl.querySelector('.btn-page-prev');
    const btnNext = paginationEl.querySelector('.btn-page-next');

    // Update page info text
    if (pageInfo) {
      pageInfo.textContent = `Página ${currentPage} de ${totalPages || 1}`;
    }

    // Enable/disable prev button
    if (btnPrev) {
      // Clone to remove old listeners
      const newPrev = btnPrev.cloneNode(true);
      btnPrev.parentNode.replaceChild(newPrev, btnPrev);
      newPrev.disabled = currentPage <= 1;
      newPrev.addEventListener('click', () => {
        const state = this._tableState[tableId];
        if (state && state.page > 1) {
          state.page--;
          this.renderTable(tableId);
        }
      });
    }

    // Enable/disable next button
    if (btnNext) {
      // Clone to remove old listeners
      const newNext = btnNext.cloneNode(true);
      btnNext.parentNode.replaceChild(newNext, btnNext);
      newNext.disabled = currentPage >= totalPages;
      newNext.addEventListener('click', () => {
        const state = this._tableState[tableId];
        if (state && state.page < totalPages) {
          state.page++;
          this.renderTable(tableId);
        }
      });
    }
  },

  /**
   * Refresh all tables with new data (e.g., after filter change).
   * Preserves search queries but resets pagination to page 1.
   * @param {FinancialRecord[]} data - New dataset (filtered or full)
   */
  refreshTables(data) {
    const tableIds = Object.keys(this._tableConfigs);

    for (const tableId of tableIds) {
      const tableData = this.getTableData(tableId, data);
      const state = this._tableState[tableId];

      if (state) {
        state.data = tableData;
        state.page = 1;

        // Re-apply search if active
        if (state.searchQuery && state.searchQuery.trim() !== '') {
          const lowerQuery = state.searchQuery.toLowerCase();
          const columns = state.columns;

          state.filteredData = tableData.filter(row => {
            return columns.some(col => {
              const value = row[col.key];
              if (value == null) return false;

              let displayValue;
              if (col.key === 'data' && value instanceof Date) {
                displayValue = TableModule.formatDate(value);
              } else if (col.key === 'valor' || col.key === 'revenue' || col.key === 'expense' || col.key === 'netValue') {
                displayValue = KPICalculator.formatBRL(Number(value));
              } else {
                displayValue = String(value);
              }

              return displayValue.toLowerCase().includes(lowerQuery);
            });
          });
        } else {
          state.filteredData = tableData;
        }

        this.renderTable(tableId);
      }
    }
  },

  // ---- KPI Cards UI ----

  /**
   * Mapping of KPI keys to their DOM element IDs and variation element IDs.
   * @type {Array<{key: string, valueId: string, variationId: string, formatter: function}>}
   * @private
   */
  _kpiCardConfig: [
    { key: 'totalRecebido',        valueId: 'kpi-total-recebido',        variationId: 'kpi-var-total-recebido',        formatter: 'brl', colorType: 'positive' },
    { key: 'totalPago',            valueId: 'kpi-total-pago',            variationId: 'kpi-var-total-pago',            formatter: 'brl', colorType: 'negative' },
    { key: 'saldo',                valueId: 'kpi-saldo',                 variationId: 'kpi-var-saldo',                 formatter: 'brl', colorType: 'auto' },
    { key: 'recebimentosAberto',   valueId: 'kpi-recebimentos-aberto',   variationId: 'kpi-var-recebimentos-aberto',   formatter: 'brl', colorType: 'positive' },
    { key: 'pagamentosAberto',     valueId: 'kpi-pagamentos-aberto',     variationId: 'kpi-var-pagamentos-aberto',     formatter: 'brl', colorType: 'negative' },
    { key: 'fluxoMes',            valueId: 'kpi-fluxo-mes',             variationId: 'kpi-var-fluxo-mes',             formatter: 'brl', colorType: 'auto' },
    { key: 'ticketMedio',          valueId: 'kpi-ticket-medio',          variationId: 'kpi-var-ticket-medio',          formatter: 'brl', colorType: 'neutral' },
    { key: 'quantidadeLancamentos', valueId: 'kpi-quantidade-lancamentos', variationId: 'kpi-var-quantidade-lancamentos', formatter: 'integer', colorType: 'neutral' }
  ],

  /**
   * Store previous KPI values for animation transitions.
   * @type {Object<string, number>}
   * @private
   */
  _previousKPIValues: {},

  /**
   * Calculate KPIs from data and update all 8 KPI cards in the DOM.
   * Applies color classes based on getValueColor, shows variation indicators,
   * and triggers number transition animations.
   *
   * @param {FinancialRecord[]} data - Financial records to calculate KPIs from
   */
  updateKPICards(data) {
    const kpis = KPICalculator.calculateAll(data);

    for (const config of this._kpiCardConfig) {
      const valueEl = document.getElementById(config.valueId);
      const variationEl = document.getElementById(config.variationId);

      if (!valueEl) continue;

      const newValue = kpis[config.key] || 0;
      const formatter = config.formatter === 'brl'
        ? KPICalculator.formatBRL
        : KPICalculator.formatInteger;

      // Animate the value transition
      this.animateKPIValue(valueEl, newValue, formatter.bind(KPICalculator));

      // Apply color class based on KPI type
      let colorClass;
      if (config.colorType === 'positive') colorClass = 'positive';
      else if (config.colorType === 'negative') colorClass = 'negative';
      else if (config.colorType === 'auto') colorClass = newValue >= 0 ? 'positive' : 'negative';
      else colorClass = 'neutral';
      valueEl.classList.remove('positive', 'negative', 'neutral');
      valueEl.classList.add(colorClass);

      // Update variation indicator
      if (variationEl) {
        const variation = kpis.variacoes[config.key];
        if (variation !== null && variation !== undefined) {
          const sign = variation >= 0 ? '+' : '';
          const variationText = `${sign}${variation.toFixed(1)}%`;
          variationEl.textContent = variationText;

          // Apply color to variation indicator
          const varColor = KPICalculator.getValueColor(variation);
          variationEl.classList.remove('positive', 'negative', 'neutral');
          variationEl.classList.add(varColor);
          variationEl.setAttribute('aria-label', `Variação mensal: ${variationText}`);
        } else {
          variationEl.textContent = '';
          variationEl.classList.remove('positive', 'negative', 'neutral');
          variationEl.setAttribute('aria-label', 'Variação mensal: sem dados');
        }
      }

      // Store current value for next animation
      this._previousKPIValues[config.key] = newValue;
    }
  },

  /**
   * Animate a KPI value element from its current displayed value to a new value
   * over 400ms using requestAnimationFrame for smooth transitions.
   *
   * @param {HTMLElement} element - The DOM element displaying the KPI value
   * @param {number} newValue - The target numeric value
   * @param {function} formatter - Formatting function (formatBRL or formatInteger)
   */
  animateKPIValue(element, newValue, formatter) {
    if (!element) return;

    // Parse the current displayed value to get the starting number
    const currentText = element.textContent || '';
    const startValue = this._parseDisplayedValue(currentText);

    // If values are the same, just set and return
    if (startValue === newValue) {
      element.textContent = formatter(newValue);
      return;
    }

    const duration = 400; // 400ms animation
    const startTime = performance.now();
    const diff = newValue - startValue;

    // Cancel any existing animation on this element
    if (element._kpiAnimationId) {
      cancelAnimationFrame(element._kpiAnimationId);
    }

    const animate = (currentTime) => {
      const elapsed = currentTime - startTime;
      const progress = Math.min(elapsed / duration, 1);

      // Ease-out cubic for smooth deceleration
      const eased = 1 - Math.pow(1 - progress, 3);

      const currentValue = startValue + diff * eased;
      element.textContent = formatter(currentValue);

      if (progress < 1) {
        element._kpiAnimationId = requestAnimationFrame(animate);
      } else {
        // Ensure final value is exact
        element.textContent = formatter(newValue);
        element._kpiAnimationId = null;
      }
    };

    element._kpiAnimationId = requestAnimationFrame(animate);
  },

  /**
   * Parse a displayed BRL or integer value back to a number.
   * Handles formats like "R$ 1.234,56", "R$ -1.234,56", "1.234", "-1.234".
   *
   * @param {string} text - The displayed text value
   * @returns {number} Parsed numeric value, or 0 if unparseable
   * @private
   */
  _parseDisplayedValue(text) {
    if (!text || text.trim() === '') return 0;

    // Remove "R$ " prefix and whitespace
    let cleaned = text.replace(/R\$\s*/g, '').trim();

    // Handle negative sign
    const isNegative = cleaned.startsWith('-');
    if (isNegative) cleaned = cleaned.substring(1);

    // Check if it has a comma (BRL decimal separator)
    if (cleaned.includes(',')) {
      // BRL format: remove dots (thousands), replace comma with dot (decimal)
      cleaned = cleaned.replace(/\./g, '').replace(',', '.');
    } else {
      // Integer format: remove dots (thousands)
      cleaned = cleaned.replace(/\./g, '');
    }

    const value = parseFloat(cleaned);
    if (isNaN(value)) return 0;
    return isNegative ? -value : value;
  },

  /**
   * Render all KPI cards from data. Alias for updateKPICards.
   * Calculates KPIs using KPICalculator.calculateAll(data), then updates each card.
   * @param {FinancialRecord[]} data - Financial records
   */
  renderKPIs(data) {
    this.updateKPICards(data);
  },

  /**
   * Update a single KPI variation element with formatted percentage and color.
   * Displays "+12%" or "-5%" with appropriate color class.
   *
   * @param {HTMLElement} variationElement - The DOM element for the variation indicator
   * @param {number|null} percentage - The percentage variation value, or null if unavailable
   */
  updateKPIVariation(variationElement, percentage) {
    if (!variationElement) return;

    if (percentage !== null && percentage !== undefined) {
      const sign = percentage >= 0 ? '+' : '';
      const variationText = `${sign}${percentage.toFixed(1)}%`;
      variationElement.textContent = variationText;

      // Apply color class
      const colorClass = KPICalculator.getValueColor(percentage);
      variationElement.classList.remove('positive', 'negative', 'neutral');
      variationElement.classList.add(colorClass);
      variationElement.setAttribute('aria-label', `Variação mensal: ${variationText}`);
    } else {
      variationElement.textContent = '';
      variationElement.classList.remove('positive', 'negative', 'neutral');
      variationElement.setAttribute('aria-label', 'Variação mensal: sem dados');
    }
  },

  // ---- Insights UI Rendering ----

  /**
   * Render all insight cards by calling InsightsEngine.generateInsights(data)
   * and updating each insight section in the DOM.
   * Shows "insufficient data" message when fewer than 3 months of history.
   *
   * @param {FinancialRecord[]} data - Financial records to generate insights from
   */
  renderInsights(data) {
    const insights = InsightsEngine.generateInsights(data);

    // Render each insight section
    this._renderAlerts(insights.alerts);
    this._renderTrends(insights);
    this._renderForecast(insights.forecast);
    this._renderRankings({
      projects: insights.projectRanking,
      clients: insights.clientRanking,
      categories: insights.categoryRanking
    });
    this._renderMonthlyVariations(insights.monthlyVariations);
    this._renderCriticalConditions(insights.alerts.filter(a => a.severity === 'critical'));

    // Show/hide insufficient data message
    const insufficientEl = document.getElementById('insights-insufficient-data');
    if (insufficientEl) {
      // Check if we have fewer than 3 months of data (forecast is null)
      const hasEnoughData = insights.forecast !== null;
      insufficientEl.hidden = hasEnoughData;
    }
  },

  /**
   * Render overdue payment alerts in the alerts insight card.
   * Lists each alert with value and client/project info.
   *
   * @param {Array<{type: string, message: string, value: number, client?: string, project?: string, month?: string, change?: number, severity?: string}>} alerts - Alert objects
   */
  _renderAlerts(alerts) {
    const contentEl = document.getElementById('insight-alerts-content');
    if (!contentEl) return;

    // Get individual pending/overdue records directly from data
    const allData = DataLayer.getData();
    const pendingRecords = allData.filter(r => r.status === 'Pendente');

    if (pendingRecords.length === 0) {
      contentEl.innerHTML = '<p class="insight-placeholder">Nenhum lançamento pendente encontrado.</p>';
      return;
    }

    // Sort by date (oldest first - most urgent)
    pendingRecords.sort((a, b) => {
      const da = a.data instanceof Date ? a.data : new Date(a.data);
      const db = b.data instanceof Date ? b.data : new Date(b.data);
      return da - db;
    });

    // Show max 10 alerts in the card
    const toShow = pendingRecords.slice(0, 10);

    let html = `<p style="font-size:11px;color:var(--text-muted);margin-bottom:8px;">${pendingRecords.length} lançamento(s) pendente(s)</p>`;
    html += '<ul class="insight-list">';
    for (const record of toShow) {
      const dateStr = record.data instanceof Date ? TableModule.formatDate(record.data) : '';
      const valueFormatted = KPICalculator.formatBRL(record.valor);
      const colorClass = record.tipo === 'receita' ? 'positive' : 'negative';
      const desc = record.descricao || record.cliente || record.projeto || '';

      html += `<li class="insight-item alert-item alert-clickable" data-record-id="${record.id}" style="cursor:pointer;">`;
      html += `<span class="alert-icon">⚠️</span>`;
      html += `<span class="alert-text" style="flex:1;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;" title="${this._escapeHtml(desc)}">${dateStr} - ${this._escapeHtml(desc.substring(0, 40))}</span>`;
      html += `<span class="alert-value ${colorClass}">${valueFormatted}</span>`;
      html += `</li>`;
    }
    html += '</ul>';

    // "Ver Todos" button
    html += `<button id="btn-view-all-alerts" class="btn-view-all-alerts">📋 Ver Todos (${pendingRecords.length})</button>`;

    contentEl.innerHTML = html;

    // Make alerts clickable to open record detail
    contentEl.querySelectorAll('.alert-clickable').forEach(item => {
      item.addEventListener('click', () => {
        const recordId = item.dataset.recordId;
        const record = allData.find(r => r.id === recordId);
        if (record) RecordDetail.open(record);
      });
    });

    // "Ver Todos" button opens the alerts modal
    const btnViewAll = document.getElementById('btn-view-all-alerts');
    if (btnViewAll) {
      btnViewAll.addEventListener('click', () => {
        AlertsModal.open();
      });
    }
  },

  /**
   * Render trend direction and expense/revenue monthly changes
   * in the trends insight card.
   *
   * @param {InsightResult} insights - Full insights object containing trendDirection, expenseIncreases, revenueDrops
   */
  _renderTrends(insights) {
    const contentEl = document.getElementById('insight-trends-content');
    if (!contentEl) return;

    const { trendDirection, expenseIncreases, revenueDrops } = insights;

    let html = '';

    // Trend direction indicator
    html += `<div class="trend-direction" aria-label="Direção da tendência financeira">`;
    html += `<span class="trend-arrow">${trendDirection}</span> `;
    const trendLabel = trendDirection === '↑' ? 'Tendência de alta' :
                       trendDirection === '↓' ? 'Tendência de queda' : 'Tendência estável';
    const trendClass = trendDirection === '↑' ? 'positive' :
                       trendDirection === '↓' ? 'negative' : 'neutral';
    html += `<span class="trend-label ${trendClass}">${trendLabel}</span>`;
    html += `</div>`;

    // Revenue changes (drops)
    if (revenueDrops.length > 0) {
      html += `<div class="trend-section">`;
      html += `<h5 class="trend-subtitle">Quedas de Receita (&gt;5%)</h5>`;
      html += `<ul class="insight-list">`;
      for (const drop of revenueDrops.slice(0, 5)) {
        const changeFormatted = drop.change !== null ? drop.change.toFixed(1) : '0.0';
        html += `<li class="insight-item negative">`;
        html += `<span class="trend-month">${this._formatMonthLabel(drop.month)}</span>`;
        html += `<span class="trend-change negative">↓ ${Math.abs(changeFormatted)}%</span>`;
        html += `</li>`;
      }
      html += `</ul></div>`;
    }

    // Expense increases
    if (expenseIncreases.length > 0) {
      html += `<div class="trend-section">`;
      html += `<h5 class="trend-subtitle">Aumentos de Despesa (&gt;10%)</h5>`;
      html += `<ul class="insight-list">`;
      for (const increase of expenseIncreases.slice(0, 5)) {
        const changeFormatted = increase.change !== null ? increase.change.toFixed(1) : '0.0';
        html += `<li class="insight-item negative">`;
        html += `<span class="trend-month">${this._formatMonthLabel(increase.month)}</span>`;
        html += `<span class="trend-change negative">↑ ${changeFormatted}%</span>`;
        html += `</li>`;
      }
      html += `</ul></div>`;
    }

    if (revenueDrops.length === 0 && expenseIncreases.length === 0) {
      html += `<p class="insight-placeholder">Nenhuma variação significativa detectada.</p>`;
    }

    contentEl.innerHTML = html;
  },

  /**
   * Render forecast projections or "insufficient data" message
   * in the forecast insight card.
   *
   * @param {{months: Array<{month: string, projected: number}>, slope: number, intercept: number, confidence: string}|null} forecast - Forecast result or null
   */
  _renderForecast(forecast) {
    const contentEl = document.getElementById('insight-forecast-content');
    if (!contentEl) return;

    if (!forecast || !forecast.months || forecast.months.length === 0) {
      contentEl.innerHTML = '<p class="insight-placeholder">Dados insuficientes para previsão. São necessários pelo menos 3 meses de histórico.</p>';
      return;
    }

    const confidenceLabel = forecast.confidence === 'high' ? 'Alta' :
                            forecast.confidence === 'medium' ? 'Média' : 'Baixa';
    const confidenceClass = forecast.confidence === 'high' ? 'positive' :
                            forecast.confidence === 'medium' ? 'neutral' : 'negative';

    let html = `<div class="forecast-confidence">`;
    html += `<span>Confiança: </span><span class="${confidenceClass}">${confidenceLabel}</span>`;
    html += `</div>`;

    html += `<ul class="insight-list forecast-list" aria-label="Projeção para os próximos 3 meses">`;
    for (const entry of forecast.months) {
      const valueFormatted = KPICalculator.formatBRL(entry.projected);
      const valueClass = entry.projected >= 0 ? 'positive' : 'negative';
      html += `<li class="insight-item forecast-item">`;
      html += `<span class="forecast-month">${this._formatMonthLabel(entry.month)}</span>`;
      html += `<span class="forecast-value ${valueClass}">${valueFormatted}</span>`;
      html += `</li>`;
    }
    html += `</ul>`;

    contentEl.innerHTML = html;
  },

  /**
   * Populate the 3 ranking lists (projetos, clientes, categorias)
   * with top 10 entries each.
   *
   * @param {{projects: Array<{name: string, netValue: number}>, clients: Array<{name: string, netValue: number}>, categories: Array<{name: string, netValue: number}>}} rankings - Rankings data
   */
  _renderRankings(rankings) {
    // Render project rankings
    const projetosEl = document.getElementById('ranking-projetos');
    if (projetosEl) {
      projetosEl.innerHTML = '';
      if (rankings.projects && rankings.projects.length > 0) {
        for (const entry of rankings.projects.slice(0, 10)) {
          const li = document.createElement('li');
          const valueClass = entry.netValue >= 0 ? 'positive' : 'negative';
          li.innerHTML = `<span class="ranking-name">${this._escapeHtml(entry.name)}</span> <span class="ranking-value ${valueClass}">${KPICalculator.formatBRL(entry.netValue)}</span>`;
          projetosEl.appendChild(li);
        }
      } else {
        projetosEl.innerHTML = '<li class="insight-placeholder">Sem dados de projetos.</li>';
      }
    }

    // Render client rankings
    const clientesEl = document.getElementById('ranking-clientes');
    if (clientesEl) {
      clientesEl.innerHTML = '';
      if (rankings.clients && rankings.clients.length > 0) {
        for (const entry of rankings.clients.slice(0, 10)) {
          const li = document.createElement('li');
          const valueClass = entry.netValue >= 0 ? 'positive' : 'negative';
          li.innerHTML = `<span class="ranking-name">${this._escapeHtml(entry.name)}</span> <span class="ranking-value ${valueClass}">${KPICalculator.formatBRL(entry.netValue)}</span>`;
          clientesEl.appendChild(li);
        }
      } else {
        clientesEl.innerHTML = '<li class="insight-placeholder">Sem dados de clientes.</li>';
      }
    }

    // Render category rankings
    const categoriasEl = document.getElementById('ranking-categorias');
    if (categoriasEl) {
      categoriasEl.innerHTML = '';
      if (rankings.categories && rankings.categories.length > 0) {
        for (const entry of rankings.categories.slice(0, 10)) {
          const li = document.createElement('li');
          const valueClass = entry.netValue >= 0 ? 'positive' : 'negative';
          li.innerHTML = `<span class="ranking-name">${this._escapeHtml(entry.name)}</span> <span class="ranking-value ${valueClass}">${KPICalculator.formatBRL(entry.netValue)}</span>`;
          categoriasEl.appendChild(li);
        }
      } else {
        categoriasEl.innerHTML = '<li class="insight-placeholder">Sem dados de categorias.</li>';
      }
    }
  },

  /**
   * Render monthly variations with color coding in the monthly insight card.
   * Shows month-over-month changes for both revenue and expenses.
   *
   * @param {Array<{month: string, value: number, previousValue: number|null, change: number|null}>} variations - Monthly variation entries
   */
  _renderMonthlyVariations(variations) {
    const contentEl = document.getElementById('insight-monthly-content');
    if (!contentEl) return;

    if (!variations || variations.length === 0) {
      contentEl.innerHTML = '<p class="insight-placeholder">Sem dados de variações mensais.</p>';
      return;
    }

    // Group by month and show the most recent 6 months
    const monthMap = {};
    for (const v of variations) {
      if (!monthMap[v.month]) {
        monthMap[v.month] = { month: v.month, changes: [] };
      }
      monthMap[v.month].changes.push(v);
    }

    const sortedMonths = Object.keys(monthMap).sort().slice(-6);

    let html = '<ul class="insight-list monthly-list" aria-label="Variações mensais">';
    for (const monthKey of sortedMonths) {
      const entry = monthMap[monthKey];
      html += `<li class="insight-item monthly-item">`;
      html += `<span class="monthly-month">${this._formatMonthLabel(monthKey)}</span>`;

      for (const change of entry.changes) {
        if (change.change !== null) {
          const sign = change.change >= 0 ? '+' : '';
          const colorClass = change.change > 0 ? 'negative' : change.change < 0 ? 'positive' : 'neutral';
          // For expenses: increase is bad (negative color), decrease is good (positive color)
          // For revenue: increase is good (positive color), decrease is bad (negative color)
          // Since we mix both, use the raw value direction
          const displayClass = change.change >= 0 ? (change.value > (change.previousValue || 0) ? 'neutral' : 'positive') : 'negative';
          html += `<span class="monthly-change ${change.change > 5 ? 'negative' : change.change < -5 ? 'negative' : 'neutral'}">${sign}${change.change.toFixed(1)}%</span> `;
        }
      }

      html += `</li>`;
    }
    html += '</ul>';

    contentEl.innerHTML = html;
  },

  /**
   * Render critical conditions with warning styling.
   * Adds 'critical' class to the insight-critical card and displays
   * critical alerts with warning icon and negative-value color.
   *
   * @param {Array<{month: string, type: string, change: number, message: string, severity: string}>} criticals - Critical condition alerts
   */
  _renderCriticalConditions(criticals) {
    const contentEl = document.getElementById('insight-critical-content');
    const cardEl = document.getElementById('insight-critical');
    if (!contentEl) return;

    if (!criticals || criticals.length === 0) {
      contentEl.innerHTML = '<p class="insight-placeholder">Nenhuma condição crítica detectada.</p>';
      if (cardEl) cardEl.classList.remove('critical');
      return;
    }

    // Add critical class to the card for warning styling
    if (cardEl) cardEl.classList.add('critical');

    let html = '<ul class="insight-list critical-list" aria-label="Condições financeiras críticas">';
    for (const alert of criticals) {
      html += `<li class="insight-item critical-item">`;
      html += `<span class="critical-icon" aria-hidden="true">⚠️</span> `;
      html += `<span class="critical-text negative">${this._escapeHtml(alert.message)}</span>`;
      html += `</li>`;
    }
    html += '</ul>';

    contentEl.innerHTML = html;
  },

  /**
   * Format a month key (YYYY-MM) into a readable label (e.g., "Jan/2024").
   * @param {string} monthKey - Month in YYYY-MM format
   * @returns {string} Formatted month label
   * @private
   */
  _formatMonthLabel(monthKey) {
    if (!monthKey || typeof monthKey !== 'string') return '';
    const parts = monthKey.split('-');
    if (parts.length < 2) return monthKey;

    const monthNames = ['Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez'];
    const monthIndex = parseInt(parts[1], 10) - 1;
    const year = parts[0];

    if (monthIndex >= 0 && monthIndex < 12) {
      return `${monthNames[monthIndex]}/${year}`;
    }
    return monthKey;
  },

  /**
   * Escape HTML special characters to prevent XSS.
   * @param {string} str - Raw string
   * @returns {string} HTML-safe string
   * @private
   */
  _escapeHtml(str) {
    if (!str) return '';
    return String(str)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#039;');
  },

  // ---- File Upload UI ----

  /**
   * Maximum file size in bytes before showing confirmation dialog (50MB).
   * @type {number}
   * @private
   */
  _MAX_FILE_SIZE: 50 * 1024 * 1024,

  /**
   * Allowed file extensions for upload.
   * @type {string[]}
   * @private
   */
  _ALLOWED_EXTENSIONS: ['.xlsx', '.xls'],

  /**
   * Initialize the file upload area with drag-and-drop and click-to-open handlers.
   * Sets up event listeners on #upload-dropzone, .upload-button, and #file-input.
   */
  initUpload() {
    const dropzone = document.getElementById('upload-dropzone');
    const fileInput = document.getElementById('file-input');
    const uploadButton = dropzone ? dropzone.querySelector('.upload-button') : null;

    if (!dropzone || !fileInput) return;

    // Drag-and-drop handlers
    dropzone.addEventListener('dragover', (e) => {
      e.preventDefault();
      e.stopPropagation();
      dropzone.classList.add('dragover');
    });

    dropzone.addEventListener('dragenter', (e) => {
      e.preventDefault();
      e.stopPropagation();
      dropzone.classList.add('dragover');
    });

    dropzone.addEventListener('dragleave', (e) => {
      e.preventDefault();
      e.stopPropagation();
      dropzone.classList.remove('dragover');
    });

    dropzone.addEventListener('drop', (e) => {
      e.preventDefault();
      e.stopPropagation();
      dropzone.classList.remove('dragover');

      const files = e.dataTransfer.files;
      this._handleFiles(files);
    });

    // Click to open file picker
    dropzone.addEventListener('click', (e) => {
      // Avoid triggering if the button itself was clicked (it has its own handler)
      if (e.target === uploadButton) return;
      fileInput.click();
    });

    // Keyboard support for dropzone (Enter/Space opens file picker)
    dropzone.addEventListener('keydown', (e) => {
      if (e.key === 'Enter' || e.key === ' ') {
        e.preventDefault();
        fileInput.click();
      }
    });

    // Upload button click
    if (uploadButton) {
      uploadButton.addEventListener('click', (e) => {
        e.stopPropagation();
        fileInput.click();
      });
    }

    // File input change handler
    fileInput.addEventListener('change', () => {
      if (fileInput.files && fileInput.files.length > 0) {
        this._handleFiles(fileInput.files);
      }
      // Reset input so the same file can be re-selected
      fileInput.value = '';
    });
  },

  /**
   * Handle file(s) from drop or file picker.
   * Validates that only a single file is provided, then runs full validation.
   * @param {FileList} files - The files from the drop event or file input
   * @private
   */
  _handleFiles(files) {
    // Clear any previous error
    this._hideUploadError();

    // Reject multiple files
    if (files.length > 1) {
      this._showUploadError('Apenas um arquivo pode ser importado por vez. Por favor, selecione um único arquivo .xlsx ou .xls.');
      return;
    }

    if (files.length === 0) return;

    const file = files[0];
    this._validateAndProcess(file);
  },

  /**
   * Validate a file's extension and size, then process it.
   * Shows confirmation dialog for files >50MB.
   * @param {File} file - The file to validate
   * @private
   */
  async _validateAndProcess(file) {
    // Validate file extension
    const fileName = file.name || '';
    const extension = fileName.substring(fileName.lastIndexOf('.')).toLowerCase();

    if (!this._ALLOWED_EXTENSIONS.includes(extension)) {
      this._showUploadError(
        `Formato de arquivo inválido: "${extension}". Apenas arquivos .xlsx e .xls são aceitos. ` +
        'Verifique se o arquivo foi exportado corretamente do DOit ERP.'
      );
      return;
    }

    // Check file size - show confirmation for >50MB
    if (file.size > this._MAX_FILE_SIZE) {
      const confirmed = await this._showLargeFileDialog(file);
      if (!confirmed) return;
    }

    // Process the file
    this._processFile(file);
  },

  /**
   * Show the large file confirmation dialog and wait for user response.
   * @param {File} file - The large file
   * @returns {Promise<boolean>} True if user confirms, false if cancelled
   * @private
   */
  _showLargeFileDialog(file) {
    return new Promise((resolve) => {
      const dialog = document.getElementById('large-file-dialog');
      const message = document.getElementById('large-file-message');
      const btnConfirm = document.getElementById('btn-large-file-confirm');
      const btnCancel = document.getElementById('btn-large-file-cancel');

      if (!dialog || !btnConfirm || !btnCancel) {
        // If dialog elements don't exist, proceed anyway
        resolve(true);
        return;
      }

      // Set message with file size info
      const sizeMB = (file.size / (1024 * 1024)).toFixed(1);
      if (message) {
        message.textContent = `O arquivo "${file.name}" possui ${sizeMB} MB. ` +
          'Arquivos grandes podem levar mais tempo para processar. Deseja continuar?';
      }

      // Show dialog
      dialog.hidden = false;

      // Cleanup function to remove listeners and hide dialog
      const cleanup = () => {
        btnConfirm.removeEventListener('click', onConfirm);
        btnCancel.removeEventListener('click', onCancel);
        dialog.hidden = true;
      };

      const onConfirm = () => {
        cleanup();
        resolve(true);
      };

      const onCancel = () => {
        cleanup();
        resolve(false);
      };

      btnConfirm.addEventListener('click', onConfirm);
      btnCancel.addEventListener('click', onCancel);
    });
  },

  /**
   * Process a validated file: show loading overlay, parse with ExcelParser,
   * and transition to dashboard view on success.
   * @param {File} file - The validated Excel file to process
   * @private
   */
  async _processFile(file) {
    // Show loading overlay
    this._showLoadingOverlay('reading', 0);

    try {
      // Parse file with progress callback
      const result = await ExcelParser.parseFile(file, (stage, percent) => {
        this._updateLoadingProgress(stage, percent);
      });

      // Final progress update
      this._updateLoadingProgress('rendering', 100);

      // Short delay to show 100% before hiding
      await new Promise(r => setTimeout(r, 300));

      // Hide loading overlay
      this._hideLoadingOverlay();

      // Transition to dashboard view FIRST (so sections are visible for Chart.js)
      this._showDashboard();

      // Now trigger data:loaded event (charts need visible canvas to render)
      // Note: DataLayer.setData was already called inside parseFile,
      // but we need to re-emit for charts since sections are now visible
      requestAnimationFrame(() => {
        ChartEngine.initCharts(DataLayer.getData());
      });

    } catch (error) {
      // Hide loading overlay
      this._hideLoadingOverlay();

      // Determine error stage and show appropriate message
      const errorMessage = error.message || 'Erro desconhecido ao processar o arquivo.';
      this._showUploadError(
        `Erro ao processar o arquivo: ${errorMessage} ` +
        'Verifique se o arquivo é um Excel válido exportado do DOit ERP e tente novamente.'
      );
    }
  },

  /**
   * Show the loading overlay with initial stage and progress.
   * @param {string} stage - Current processing stage ID
   * @param {number} percent - Progress percentage (0-100)
   * @private
   */
  _showLoadingOverlay(stage, percent) {
    const overlay = document.getElementById('loading-overlay');

    if (overlay) {
      overlay.hidden = false;
      overlay.style.display = 'flex';
    }

    this._updateLoadingProgress(stage, percent);
  },

  /**
   * Update the loading overlay progress bar and stage text.
   * @param {string} stage - Current processing stage ID
   * @param {number} percent - Progress percentage (0-100)
   * @private
   */
  _updateLoadingProgress(stage, percent) {
    const message = document.getElementById('loading-message');
    const progressFill = document.getElementById('loading-progress-fill');
    const stageEl = document.getElementById('loading-stage');
    const progressBar = progressFill ? progressFill.parentElement : null;

    // Map stage to label
    const stageLabels = {
      reading: 'Lendo arquivo...',
      mapping: 'Mapeando colunas...',
      processing: 'Processando dados...',
      rendering: 'Renderizando dashboard...'
    };

    if (message) {
      message.textContent = stageLabels[stage] || 'Processando...';
    }

    if (progressFill) {
      const clampedPercent = Math.max(0, Math.min(100, percent));
      progressFill.style.width = `${clampedPercent}%`;
    }

    if (progressBar) {
      progressBar.setAttribute('aria-valuenow', String(Math.round(percent)));
    }

    if (stageEl) {
      // Show stage number (e.g., "Etapa 2 de 4")
      const stageOrder = ['reading', 'mapping', 'processing', 'rendering'];
      const stageIndex = stageOrder.indexOf(stage);
      if (stageIndex >= 0) {
        stageEl.textContent = `Etapa ${stageIndex + 1} de ${stageOrder.length}`;
      }
    }
  },

  /**
   * Hide the loading overlay.
   * @private
   */
  _hideLoadingOverlay() {
    const overlay = document.getElementById('loading-overlay');
    if (overlay) {
      overlay.hidden = true;
      overlay.style.display = 'none';
    }
  },

  /**
   * Show an error message in the upload error area.
   * @param {string} message - Error message to display
   * @private
   */
  _showUploadError(message) {
    const errorEl = document.getElementById('upload-error');
    if (errorEl) {
      errorEl.textContent = message;
      errorEl.hidden = false;
    }
  },

  /**
   * Hide the upload error message.
   * @private
   */
  _hideUploadError() {
    const errorEl = document.getElementById('upload-error');
    if (errorEl) {
      errorEl.textContent = '';
      errorEl.hidden = true;
    }
  },

  /**
   * Transition from upload view to dashboard view.
   * Hides the upload section and shows filter bar + dashboard sections.
   * @private
   */
  _showDashboard() {
    const uploadArea = document.getElementById('upload-area');
    const filterBar = document.getElementById('filter-bar');
    const dashboardSections = document.querySelectorAll('.dashboard-section');

    // Hide upload area
    if (uploadArea) {
      uploadArea.hidden = true;
    }

    // Show filter bar
    if (filterBar) {
      filterBar.hidden = false;
    }

    // Show all dashboard sections
    for (const section of dashboardSections) {
      section.hidden = false;
    }
  },

  // ---- Column Mapping UI ----

  /**
   * Internal resolve/reject references for the column mapping Promise.
   * @type {{resolve: Function|null, reject: Function|null}}
   * @private
   */
  _mappingPromise: { resolve: null, reject: null },

  /**
   * Show the manual column mapping modal.
   * Populates #mapping-fields with a row for each missing required field,
   * each containing a label and a dropdown with all unmapped column names.
   * Returns a Promise that resolves with the mapping when user confirms,
   * or rejects when user cancels.
   *
   * @param {string[]} unmappedColumns - All file columns that were not auto-mapped
   * @param {string[]} missingFields - Required fields that could not be auto-identified (e.g., ['valor', 'tipo', 'data'])
   * @returns {Promise<Object<string, string>>} Resolves with mapping object {fieldName: columnName}
   */
  showColumnMapping(unmappedColumns, missingFields) {
    return new Promise((resolve, reject) => {
      this._mappingPromise = { resolve, reject };

      const modal = document.getElementById('column-mapping-modal');
      const fieldsContainer = document.getElementById('mapping-fields');

      if (!modal || !fieldsContainer) {
        reject(new Error('Elementos do modal de mapeamento não encontrados.'));
        return;
      }

      // Clear previous content
      fieldsContainer.innerHTML = '';

      // Remove any previous inline error
      const existingError = modal.querySelector('.mapping-error');
      if (existingError) existingError.remove();

      // Field display labels (Portuguese)
      const fieldLabels = {
        valor: 'Valor',
        tipo: 'Tipo',
        data: 'Data',
        projeto: 'Projeto',
        cliente: 'Cliente',
        categoria: 'Categoria',
        status: 'Status',
        centroCusto: 'Centro de Custo',
        responsavel: 'Responsável',
        departamento: 'Departamento',
        conta: 'Conta'
      };

      // Create a mapping row for each missing required field
      for (const field of missingFields) {
        const row = document.createElement('div');
        row.className = 'mapping-row';
        row.setAttribute('data-field', field);

        const label = document.createElement('label');
        label.className = 'mapping-label';
        label.textContent = fieldLabels[field] || field;
        label.setAttribute('for', `mapping-select-${field}`);

        const select = document.createElement('select');
        select.className = 'mapping-select';
        select.id = `mapping-select-${field}`;
        select.setAttribute('aria-label', `Selecionar coluna para ${fieldLabels[field] || field}`);

        // Default empty option
        const defaultOption = document.createElement('option');
        defaultOption.value = '';
        defaultOption.textContent = '-- Selecione uma coluna --';
        select.appendChild(defaultOption);

        // Add all unmapped columns as options
        for (const col of unmappedColumns) {
          const option = document.createElement('option');
          option.value = col;
          option.textContent = col;
          select.appendChild(option);
        }

        row.appendChild(label);
        row.appendChild(select);
        fieldsContainer.appendChild(row);
      }

      // Show the modal
      modal.removeAttribute('hidden');

      // Wire confirm and cancel buttons
      this._wireColumnMappingButtons();
    });
  },

  /**
   * Hide the column mapping modal.
   * Adds the 'hidden' attribute back to #column-mapping-modal.
   */
  hideColumnMapping() {
    const modal = document.getElementById('column-mapping-modal');
    if (modal) {
      modal.setAttribute('hidden', '');
    }
  },

  /**
   * Read selected values from all mapping dropdowns and return a mapping object.
   * Maps field names to the selected column names.
   * @returns {Object<string, string>} Mapping object {fieldName: selectedColumnName}
   */
  getManualMapping() {
    const mapping = {};
    const fieldsContainer = document.getElementById('mapping-fields');
    if (!fieldsContainer) return mapping;

    const rows = fieldsContainer.querySelectorAll('.mapping-row');
    for (const row of rows) {
      const field = row.getAttribute('data-field');
      const select = row.querySelector('.mapping-select');
      if (field && select && select.value) {
        mapping[field] = select.value;
      }
    }

    return mapping;
  },

  /**
   * Wire #btn-mapping-confirm and #btn-mapping-cancel event handlers.
   * Confirm validates all required fields are mapped before resolving.
   * Cancel hides modal and rejects the promise.
   * @private
   */
  _wireColumnMappingButtons() {
    const btnConfirm = document.getElementById('btn-mapping-confirm');
    const btnCancel = document.getElementById('btn-mapping-cancel');

    // Clone buttons to remove any previous listeners
    if (btnConfirm) {
      const newConfirm = btnConfirm.cloneNode(true);
      btnConfirm.parentNode.replaceChild(newConfirm, btnConfirm);

      newConfirm.addEventListener('click', () => {
        const mapping = this.getManualMapping();

        // Validate that all required fields are mapped
        const requiredFields = Object.entries(COLUMN_DEFINITIONS)
          .filter(([, def]) => def.required)
          .map(([fieldName]) => fieldName);

        const fieldsContainer = document.getElementById('mapping-fields');
        const missingFieldsInModal = [];

        // Only validate fields that are actually shown in the modal
        if (fieldsContainer) {
          const rows = fieldsContainer.querySelectorAll('.mapping-row');
          for (const row of rows) {
            const field = row.getAttribute('data-field');
            if (requiredFields.includes(field) && !mapping[field]) {
              missingFieldsInModal.push(field);
            }
          }
        }

        if (missingFieldsInModal.length > 0) {
          // Show inline error
          this._showMappingError('Todos os campos obrigatórios (Valor, Tipo, Data) devem ser mapeados.');
          return;
        }

        // Valid mapping - hide modal and resolve
        this.hideColumnMapping();
        if (this._mappingPromise.resolve) {
          this._mappingPromise.resolve(mapping);
          this._mappingPromise = { resolve: null, reject: null };
        }
      });
    }

    if (btnCancel) {
      const newCancel = btnCancel.cloneNode(true);
      btnCancel.parentNode.replaceChild(newCancel, btnCancel);

      newCancel.addEventListener('click', () => {
        this.hideColumnMapping();

        // Show upload area again
        const uploadArea = document.getElementById('upload-area');
        if (uploadArea) {
          uploadArea.removeAttribute('hidden');
          uploadArea.style.display = '';
        }

        if (this._mappingPromise.reject) {
          this._mappingPromise.reject(new Error('Mapeamento cancelado pelo usuário.'));
          this._mappingPromise = { resolve: null, reject: null };
        }
      });
    }
  },

  /**
   * Show an inline error message inside the column mapping modal.
   * @param {string} message - Error message to display
   * @private
   */
  _showMappingError(message) {
    const modal = document.getElementById('column-mapping-modal');
    if (!modal) return;

    // Remove existing error if any
    const existingError = modal.querySelector('.mapping-error');
    if (existingError) existingError.remove();

    const errorEl = document.createElement('p');
    errorEl.className = 'mapping-error';
    errorEl.setAttribute('role', 'alert');
    errorEl.setAttribute('aria-live', 'assertive');
    errorEl.textContent = message;
    errorEl.style.cssText = 'color: #dc3545; font-size: 14px; margin-top: 8px; font-weight: 500;';

    // Insert before the modal-actions div
    const actionsDiv = modal.querySelector('.modal-actions');
    if (actionsDiv) {
      actionsDiv.parentNode.insertBefore(errorEl, actionsDiv);
    } else {
      modal.querySelector('.modal-content').appendChild(errorEl);
    }
  },

  // ---- Sidebar Navigation ----

  /**
   * Whether the sidebar is currently in mobile overlay mode.
   * @type {boolean}
   * @private
   */
  _sidebarOverlayOpen: false,

  /**
   * Reference to the outside-click handler for cleanup.
   * @type {Function|null}
   * @private
   */
  _outsideClickHandler: null,

  /**
   * IntersectionObserver instance for scroll spy.
   * @type {IntersectionObserver|null}
   * @private
   */
  _scrollSpyObserver: null,

  /**
   * Initialize sidebar navigation: click handlers on menu items,
   * hamburger toggle, outside-click close, and scroll spy.
   * Called once during app initialization.
   */
  initSidebar() {
    const sidebar = document.getElementById('sidebar');
    const toggle = document.getElementById('sidebar-toggle');
    const sidebarItems = document.querySelectorAll('.sidebar-item');

    if (!sidebar) return;

    // 1. Click handlers on each .sidebar-item
    sidebarItems.forEach(item => {
      item.addEventListener('click', (e) => {
        e.preventDefault();
        const sectionId = item.getAttribute('data-section');
        if (!sectionId) return;

        // Smooth scroll to target section
        const targetSection = document.getElementById(sectionId);
        if (targetSection) {
          targetSection.scrollIntoView({ behavior: 'smooth' });
        }

        // Set active class on clicked item, remove from others
        this._setActiveMenuItem(item);

        // Emit 'section:navigated' event
        EventBus.emit('section:navigated', { sectionId });

        // If sidebar is in overlay mode (mobile), close it
        if (this._sidebarOverlayOpen) {
          this._closeSidebarOverlay();
        }
      });
    });

    // 2. Hamburger toggle (#sidebar-toggle)
    if (toggle) {
      toggle.addEventListener('click', () => {
        this._toggleSidebar();
      });
    }

    // 3. Initialize scroll spy with IntersectionObserver
    this._initScrollSpy();
  },

  /**
   * Set the active class on the given menu item and remove it from all others.
   * Also updates aria-current attribute for accessibility.
   * @param {HTMLElement} activeItem - The menu item to mark as active
   * @private
   */
  _setActiveMenuItem(activeItem) {
    const allItems = document.querySelectorAll('.sidebar-item');
    allItems.forEach(item => {
      item.classList.remove('active');
      item.removeAttribute('aria-current');
    });
    activeItem.classList.add('active');
    activeItem.setAttribute('aria-current', 'page');
  },

  /**
   * Toggle the sidebar open/closed state.
   * Adds/removes 'open' class on #sidebar and updates aria-expanded on toggle button.
   * When opening, attaches an outside-click listener to close the overlay.
   * @private
   */
  _toggleSidebar() {
    const sidebar = document.getElementById('sidebar');
    const toggle = document.getElementById('sidebar-toggle');
    if (!sidebar) return;

    const isOpen = sidebar.classList.contains('open');

    if (isOpen) {
      this._closeSidebarOverlay();
    } else {
      this._openSidebarOverlay();
    }

    // Update aria-expanded
    if (toggle) {
      toggle.setAttribute('aria-expanded', String(!isOpen));
    }
  },

  /**
   * Open the sidebar as an overlay (mobile mode).
   * Adds 'open' class and attaches outside-click listener.
   * @private
   */
  _openSidebarOverlay() {
    const sidebar = document.getElementById('sidebar');
    if (!sidebar) return;

    sidebar.classList.add('open');
    this._sidebarOverlayOpen = true;

    // Attach outside-click handler (delayed to avoid catching the toggle click itself)
    setTimeout(() => {
      this._outsideClickHandler = (e) => {
        const sidebar = document.getElementById('sidebar');
        const toggle = document.getElementById('sidebar-toggle');
        if (!sidebar) return;

        // Close if click is outside sidebar and not on the toggle button
        if (!sidebar.contains(e.target) && (!toggle || !toggle.contains(e.target))) {
          this._closeSidebarOverlay();
        }
      };
      document.addEventListener('click', this._outsideClickHandler);
    }, 0);
  },

  /**
   * Close the sidebar overlay.
   * Removes 'open' class and detaches outside-click listener.
   * @private
   */
  _closeSidebarOverlay() {
    const sidebar = document.getElementById('sidebar');
    const toggle = document.getElementById('sidebar-toggle');
    if (!sidebar) return;

    sidebar.classList.remove('open');
    this._sidebarOverlayOpen = false;

    // Update aria-expanded
    if (toggle) {
      toggle.setAttribute('aria-expanded', 'false');
    }

    // Remove outside-click handler
    if (this._outsideClickHandler) {
      document.removeEventListener('click', this._outsideClickHandler);
      this._outsideClickHandler = null;
    }
  },

  /**
   * Initialize IntersectionObserver-based scroll spy.
   * Updates the active menu item based on which section is currently
   * most visible in the viewport.
   * @private
   */
  _initScrollSpy() {
    const sections = document.querySelectorAll('.dashboard-section');
    if (sections.length === 0) return;

    // Track which sections are currently intersecting
    const visibleSections = new Map();

    const observerOptions = {
      root: null,
      rootMargin: '-10% 0px -60% 0px',
      threshold: 0
    };

    this._scrollSpyObserver = new IntersectionObserver((entries) => {
      entries.forEach(entry => {
        if (entry.isIntersecting) {
          visibleSections.set(entry.target.id, entry.intersectionRatio);
        } else {
          visibleSections.delete(entry.target.id);
        }
      });

      // Find the topmost visible section
      if (visibleSections.size > 0) {
        // Get the first visible section in DOM order
        let activeSectionId = null;
        for (const section of sections) {
          if (visibleSections.has(section.id)) {
            activeSectionId = section.id;
            break;
          }
        }

        if (activeSectionId) {
          const menuItem = document.querySelector(`.sidebar-item[data-section="${activeSectionId}"]`);
          if (menuItem && !menuItem.classList.contains('active')) {
            this._setActiveMenuItem(menuItem);
          }
        }
      }
    }, observerOptions);

    // Observe all dashboard sections
    sections.forEach(section => {
      this._scrollSpyObserver.observe(section);
    });
  },

  // ---- Error Handling & User Feedback ----

  /**
   * Error type to user-friendly message mapping (Portuguese).
   * @type {Object<string, string>}
   * @private
   */
  _errorMessages: {
    'invalid-format': 'O arquivo selecionado não é um formato Excel válido (.xlsx ou .xls).',
    'corrupted-file': 'O arquivo está corrompido ou não pode ser lido. Tente exportar novamente do DOit ERP.',
    'missing-columns': 'Colunas obrigatórias não encontradas no arquivo.',
    'no-financial-data': 'Nenhuma planilha com dados financeiros reconhecíveis foi encontrada.',
    'file-too-large': 'O arquivo excede o tamanho máximo suportado (50MB).',
    'browser-incompatible': 'Seu navegador não suporta os recursos necessários para esta aplicação.'
  },

  /**
   * Processing stage labels in Portuguese.
   * @type {Object<string, string>}
   * @private
   */
  _stageLabels: {
    'reading': 'Leitura do arquivo',
    'mapping': 'Mapeamento de colunas',
    'processing': 'Processamento de dados',
    'rendering': 'Renderização do dashboard'
  },

  /**
   * Display an error message in the #upload-error element.
   * Shows error icon, message text, processing stage info, and suggested action.
   * Error remains visible until dismissed via close button or a new upload is initiated.
   *
   * @param {string} message - The error message to display
   * @param {string} [stage] - The processing stage where the error occurred (reading, mapping, processing, rendering)
   * @param {string} [action] - Suggested corrective action for the user
   */
  showError(message, stage, action) {
    const errorEl = document.getElementById('upload-error');
    const uploadArea = document.getElementById('upload-area');
    if (!errorEl) return;

    // Build error content
    const stageText = stage ? (this._stageLabels[stage] || stage) : '';
    const actionText = action || '';

    let html = '<div class="error-content">';
    html += '<span class="error-icon" aria-hidden="true">&#9888;</span>';
    html += '<div class="error-body">';
    html += '<p class="error-message">' + this._escapeHtml(message) + '</p>';
    if (stageText) {
      html += '<p class="error-stage"><strong>Etapa:</strong> ' + this._escapeHtml(stageText) + '</p>';
    }
    if (actionText) {
      html += '<p class="error-action"><strong>Sugestão:</strong> ' + this._escapeHtml(actionText) + '</p>';
    }
    html += '</div>';
    html += '<button class="error-close-btn" aria-label="Fechar mensagem de erro" title="Fechar">&#10005;</button>';
    html += '</div>';

    errorEl.innerHTML = html;
    errorEl.hidden = false;

    // Add 'error' class to upload area
    if (uploadArea) {
      uploadArea.classList.add('error');
    }

    // Attach close button handler
    const closeBtn = errorEl.querySelector('.error-close-btn');
    if (closeBtn) {
      closeBtn.addEventListener('click', () => this.hideError());
    }
  },

  /**
   * Hide the error message and remove the 'error' class from the upload area.
   */
  hideError() {
    const errorEl = document.getElementById('upload-error');
    const uploadArea = document.getElementById('upload-area');

    if (errorEl) {
      errorEl.hidden = true;
      errorEl.innerHTML = '';
    }
    if (uploadArea) {
      uploadArea.classList.remove('error');
    }
  },

  /**
   * Show a non-modal message listing expected columns when no mappable columns are found.
   * Displays the list of expected column names and allows the user to upload a different file.
   *
   * @param {string[]} [expectedColumns] - List of expected column names to display
   */
  showExpectedColumnsMessage(expectedColumns) {
    const columns = expectedColumns || Object.keys(COLUMN_DEFINITIONS).map(key => {
      const aliases = COLUMN_DEFINITIONS[key].aliases;
      return aliases[0].charAt(0).toUpperCase() + aliases[0].slice(1);
    });

    const columnList = columns.join(', ');
    const message = 'Nenhuma coluna reconhecível foi encontrada no arquivo. Colunas esperadas: ' + columnList + '.';
    const action = 'Verifique se o arquivo foi exportado corretamente do DOit ERP e tente novamente.';

    this.showError(message, 'mapping', action);
  },

  /**
   * Check browser compatibility for required features: FileReader, Drag and Drop, Blob.
   * If any feature is missing, shows a compatibility warning and disables upload.
   *
   * @returns {boolean} true if browser is compatible, false otherwise
   */
  checkBrowserCompatibility() {
    const missingFeatures = [];

    // Check FileReader support
    if (typeof FileReader === 'undefined') {
      missingFeatures.push('FileReader');
    }

    // Check Drag and Drop support
    if (typeof DataTransfer === 'undefined' && typeof window.DragEvent === 'undefined') {
      missingFeatures.push('Drag and Drop');
    }

    // Check Blob support
    if (typeof Blob === 'undefined') {
      missingFeatures.push('Blob');
    }

    if (missingFeatures.length > 0) {
      const message = 'Seu navegador não suporta os seguintes recursos necessários: ' + missingFeatures.join(', ') + '.';
      const action = 'Utilize um navegador atualizado. Versões mínimas suportadas: Chrome 80+, Firefox 78+, Edge 80+.';

      this.showError(message, null, action);

      // Disable upload functionality
      const dropzone = document.getElementById('upload-dropzone');
      const fileInput = document.getElementById('file-input');
      const uploadBtn = document.querySelector('.upload-button');

      if (dropzone) {
        dropzone.setAttribute('aria-disabled', 'true');
        dropzone.style.opacity = '0.5';
        dropzone.style.pointerEvents = 'none';
      }
      if (fileInput) {
        fileInput.disabled = true;
      }
      if (uploadBtn) {
        uploadBtn.disabled = true;
      }

      return false;
    }

    return true;
  },

  /**
   * Show the loading overlay with stage message and progress bar.
   * Adds 'visible' class to the overlay element.
   *
   * @param {string} stage - Current loading stage ID (reading, mapping, processing, rendering)
   * @param {number} [progress=0] - Progress percentage (0-100)
   */
  showLoadingOverlay(stage, progress) {
    const overlay = document.getElementById('loading-overlay');
    const message = document.getElementById('loading-message');
    const progressFill = document.getElementById('loading-progress-fill');
    const stageEl = document.getElementById('loading-stage');
    const progressBar = overlay ? overlay.querySelector('.loading-progress-bar') : null;

    if (!overlay) return;

    // Find the matching stage from LOADING_STAGES
    const stageConfig = LOADING_STAGES.find(s => s.id === stage);
    const stageLabel = stageConfig ? stageConfig.label : 'Carregando...';

    // Show overlay
    overlay.hidden = false;
    overlay.style.display = 'flex';
    overlay.classList.add('visible');

    // Update message
    if (message) {
      message.textContent = stageLabel;
    }

    // Update stage text
    if (stageEl) {
      const stageIndex = LOADING_STAGES.findIndex(s => s.id === stage);
      if (stageIndex >= 0) {
        stageEl.textContent = 'Etapa ' + (stageIndex + 1) + ' de ' + LOADING_STAGES.length;
      }
    }

    // Update progress bar
    const progressValue = Math.min(Math.max(progress || 0, 0), 100);
    if (progressFill) {
      progressFill.style.width = progressValue + '%';
    }
    if (progressBar) {
      progressBar.setAttribute('aria-valuenow', String(progressValue));
    }
  },

  /**
   * Hide the loading overlay and remove 'visible' class.
   */
  hideLoadingOverlay() {
    const overlay = document.getElementById('loading-overlay');
    const progressFill = document.getElementById('loading-progress-fill');
    const progressBar = overlay ? overlay.querySelector('.loading-progress-bar') : null;

    if (!overlay) return;

    overlay.hidden = true;
    overlay.style.display = 'none';
    overlay.classList.remove('visible');

    // Reset progress
    if (progressFill) {
      progressFill.style.width = '0%';
    }
    if (progressBar) {
      progressBar.setAttribute('aria-valuenow', '0');
    }
  }
};

// ============================================================
// SECTION: Event Listeners & Initialization
// ============================================================

// ============================================================
// SECTION: Record Detail Modal
// ============================================================

const RecordDetail = {
  _fields: [
    { key: 'descricao', label: 'Descrição' },
    { key: 'data', label: 'Data', format: 'date' },
    { key: 'valor', label: 'Valor', format: 'currency' },
    { key: 'tipo', label: 'Tipo', format: 'badge-tipo' },
    { key: 'status', label: 'Status', format: 'badge-status' },
    { key: 'projeto', label: 'Projeto' },
    { key: 'idProjeto', label: 'ID Projeto' },
    { key: 'cliente', label: 'De / Para' },
    { key: 'idCliente', label: 'ID De / Para' },
    { key: 'tipoClassif', label: '1ª Categoria' },
    { key: 'nivel1', label: '2ª Categoria' },
    { key: 'nivel2', label: '3ª Categoria' },
    { key: 'tipoPagamento', label: 'Tipo Pagamento' },
    { key: 'centroCusto', label: 'Departamento' },
    { key: 'conta', label: 'Conta' },
  ],

  open(record) {
    const modal = document.getElementById('record-detail-modal');
    const body = document.getElementById('record-detail-body');
    if (!modal || !body) return;

    let html = '';
    for (const field of this._fields) {
      const raw = record[field.key];
      if (raw == null || raw === '') continue;

      const value = this._formatValue(raw, field.format, record);
      html += `<div class="record-field"><span class="record-field-label">${field.label}</span><span class="record-field-value${this._valueClass(raw, field.format, record)}">${value}</span></div>`;
    }

    body.innerHTML = html;
    modal.hidden = false;
  },

  close() {
    const modal = document.getElementById('record-detail-modal');
    if (modal) modal.hidden = true;
  },

  _formatValue(raw, format, record) {
    if (!format) return this._esc(String(raw));
    if (format === 'date') return raw instanceof Date ? TableModule.formatDate(raw) : this._esc(String(raw));
    if (format === 'currency') return KPICalculator.formatBRL(raw);
    if (format === 'badge-tipo') {
      const cls = raw === 'receita' ? 'badge-receita' : 'badge-despesa';
      const label = raw === 'receita' ? 'Receita' : 'Despesa';
      return `<span class="record-badge ${cls}">${label}</span>`;
    }
    if (format === 'badge-status') {
      const cls = raw === 'Pago' ? 'badge-pago' : 'badge-pendente';
      return `<span class="record-badge ${cls}">${this._esc(raw)}</span>`;
    }
    return this._esc(String(raw));
  },

  _valueClass(raw, format, record) {
    if (format === 'currency') return record.tipo === 'receita' ? ' value-positive' : ' value-negative';
    return '';
  },

  _esc(s) { return s ? String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;') : ''; },

  init() {
    // Close button
    const btn = document.getElementById('btn-close-record');
    if (btn) btn.addEventListener('click', () => this.close());

    // Close on overlay click
    const modal = document.getElementById('record-detail-modal');
    if (modal) modal.addEventListener('click', (e) => { if (e.target === modal) this.close(); });

    // Delegate click on table rows in .table-card
    document.addEventListener('click', (e) => {
      const row = e.target.closest('.table-card tbody tr');
      if (!row) return;

      // Aggregate rows (projetos-lucrativos, clientes-faturamento) → open popup with filtered records
      if (row.dataset.aggregateName) {
        const name = row.dataset.aggregateName;
        const tableType = row.dataset.aggregateTable;
        const allData = DataLayer.getData();
        let filtered, title;
        if (tableType === 'projetos-lucrativos') {
          if (name === '(sem projeto)') {
            filtered = allData.filter(r => !r.projeto || r.projeto === '' || r.projeto === '(sem projeto)');
          } else {
            filtered = allData.filter(r => r.projeto === name);
          }
          title = 'Lançamentos: ' + name;
        } else if (tableType === 'clientes-faturamento') {
          if (name === '(sem cliente)') {
            filtered = allData.filter(r => !r.cliente || r.cliente === '' || r.cliente === '(sem cliente)');
          } else {
            filtered = allData.filter(r => r.cliente === name);
          }
          title = 'Lançamentos: ' + name;
        }
        if (filtered && window.KPIDetailModal) {
          KPIDetailModal._kpiKey = 'custom';
          KPIDetailModal._page = 1;
          KPIDetailModal._sortColumn = 'data';
          KPIDetailModal._sortDirection = 'asc';
          KPIDetailModal._records = filtered;
          const titleEl = document.getElementById('kpi-detail-title');
          if (titleEl) titleEl.textContent = title;
          KPIDetailModal._renderSummary();
          KPIDetailModal._renderTable();
          const modal = document.getElementById('kpi-detail-modal');
          if (modal) modal.removeAttribute('hidden');
          document.body.style.overflow = 'hidden';
        }
        return;
      }

      // Individual record rows → open record detail
      const recordId = row.dataset.recordId;
      if (!recordId) return;

      const allData = DataLayer.getData();
      const record = allData.find(r => r.id === recordId);
      if (record) this.open(record);
    });
  }
};

// Initialize on DOM ready
document.addEventListener('DOMContentLoaded', () => RecordDetail.init());

// Global ESC key handler - close any visible modal
document.addEventListener('keydown', function(e) {
  if (e.key === 'Escape' || e.keyCode === 27) {
    var modals = document.querySelectorAll('.modal-overlay:not([hidden])');
    if (modals.length > 0) {
      // Close the last (topmost) open modal
      modals[modals.length - 1].setAttribute('hidden', '');
      document.body.style.overflow = '';
      e.preventDefault();
    }
  }
});

// ============================================================
// SECTION: Pagination Renderer
// ============================================================

/**
 * Compute the array of page indicators to display.
 * When totalPages ≤ 7, returns all page numbers.
 * When totalPages > 7, applies truncation: first page, last page, current ± 1, with ellipsis for gaps.
 * @param {number} currentPage - 1-indexed current page
 * @param {number} totalPages - Total number of pages
 * @returns {Array<number|string>} Array of page numbers and 'ellipsis' markers
 */
function computePageButtons(currentPage, totalPages) {
  if (totalPages <= 7) {
    const pages = [];
    for (let i = 1; i <= totalPages; i++) pages.push(i);
    return pages;
  }

  const pages = new Set();
  pages.add(1);
  pages.add(totalPages);
  pages.add(currentPage);
  if (currentPage - 1 > 0) pages.add(currentPage - 1);
  if (currentPage + 1 <= totalPages) pages.add(currentPage + 1);

  const sorted = [...pages].sort((a, b) => a - b);
  const result = [];
  for (let i = 0; i < sorted.length; i++) {
    if (i > 0 && sorted[i] - sorted[i - 1] > 1) {
      result.push('ellipsis');
    }
    result.push(sorted[i]);
  }
  return result;
}

/**
 * Render pagination HTML into a container element.
 * @param {HTMLElement} container - The pagination wrapper element
 * @param {number} currentPage - Current active page (1-indexed)
 * @param {number} totalPages - Total pages available
 * @param {function(number): void} onPageChange - Callback when page is selected
 */
function renderPagination(container, currentPage, totalPages, onPageChange) {
  if (!container) return;
  container.innerHTML = '';

  // Hide pagination if no pages
  if (totalPages <= 0) {
    container.style.display = 'none';
    return;
  }
  container.style.display = 'flex';
  container.className = 'pagination-controls';

  // First page button
  const firstBtn = document.createElement('button');
  firstBtn.className = 'page-btn';
  firstBtn.textContent = '«';
  firstBtn.title = 'Primeira página';
  firstBtn.disabled = currentPage === 1;
  firstBtn.addEventListener('click', () => onPageChange(1));
  container.appendChild(firstBtn);

  // Previous button
  const prevBtn = document.createElement('button');
  prevBtn.className = 'page-btn';
  prevBtn.textContent = '‹';
  prevBtn.title = 'Página anterior';
  prevBtn.disabled = currentPage === 1;
  prevBtn.addEventListener('click', () => onPageChange(currentPage - 1));
  container.appendChild(prevBtn);

  // Page number buttons
  const pages = computePageButtons(currentPage, totalPages);
  for (const page of pages) {
    if (page === 'ellipsis') {
      const ellipsis = document.createElement('span');
      ellipsis.className = 'pagination-ellipsis';
      ellipsis.textContent = '…';
      container.appendChild(ellipsis);
    } else {
      const btn = document.createElement('button');
      btn.className = 'page-btn' + (page === currentPage ? ' active' : '');
      btn.textContent = String(page);
      if (page === currentPage) {
        btn.setAttribute('aria-current', 'page');
      }
      btn.addEventListener('click', () => onPageChange(page));
      container.appendChild(btn);
    }
  }

  // Next button
  const nextBtn = document.createElement('button');
  nextBtn.className = 'page-btn';
  nextBtn.textContent = '›';
  nextBtn.title = 'Próxima página';
  nextBtn.disabled = currentPage === totalPages;
  nextBtn.addEventListener('click', () => onPageChange(currentPage + 1));
  container.appendChild(nextBtn);

  // Last page button
  const lastBtn = document.createElement('button');
  lastBtn.className = 'page-btn';
  lastBtn.textContent = '»';
  lastBtn.title = 'Última página';
  lastBtn.disabled = currentPage === totalPages;
  lastBtn.addEventListener('click', () => onPageChange(totalPages));
  container.appendChild(lastBtn);
}

// ============================================================
// SECTION: KPI Detail Modal
// ============================================================

const KPIDetailModal = {
  _page: 1,
  _pageSize: 20,
  _records: [],
  _kpiKey: '',
  _sortColumn: 'data',
  _sortDirection: 'asc',

  _kpiConfig: {
    totalRecebido:        { title: 'Total Recebido',         filter: r => r.tipo === 'receita' && r.status === 'Pago' },
    totalPago:            { title: 'Total Pago',             filter: r => r.tipo === 'despesa' && r.status === 'Pago' },
    saldo:                { title: 'Saldo (Recebido - Pago)', filter: r => r.status === 'Pago' },
    recebimentosAberto:   { title: 'Recebimentos em Aberto', filter: r => r.tipo === 'receita' && r.status === 'Pendente' },
    pagamentosAberto:     { title: 'Pagamentos em Aberto',   filter: r => r.tipo === 'despesa' && r.status === 'Pendente' },
    fluxoMes:            { title: 'Fluxo do Mês',           filter: r => { const now = new Date(); const d = r.data instanceof Date ? r.data : new Date(r.data); return d.getMonth() === now.getMonth() && d.getFullYear() === now.getFullYear(); } },
    ticketMedio:          { title: 'Ticket Médio',           filter: () => true },
    quantidadeLancamentos: { title: 'Quantidade de Lançamentos', filter: () => true }
  },

  open(kpiKey) {
    const config = this._kpiConfig[kpiKey];
    if (!config) return;

    this._kpiKey = kpiKey;
    this._page = 1;
    this._sortColumn = 'data';
    this._sortDirection = 'asc';

    // For quantidadeLancamentos, include all records (even excluded)
    const includeAll = (kpiKey === 'quantidadeLancamentos' || kpiKey === 'ticketMedio');
    const data = includeAll ? DataLayer.getData(true) : (UIController._currentData || DataLayer.getData());
    this._records = data.filter(config.filter);

    // Title
    const titleEl = document.getElementById('kpi-detail-title');
    if (titleEl) titleEl.textContent = config.title;

    // Summary
    this._renderSummary();
    this._renderTable();

    const modal = document.getElementById('kpi-detail-modal');
    if (modal) {
      modal.removeAttribute('hidden');
      modal.style.display = 'flex';
    }
    document.body.style.overflow = 'hidden';
  },

  // Called from inline script when modal is already visible
  _renderFromKey(kpiKey) {
    const config = this._kpiConfig[kpiKey];
    if (!config) return;

    this._kpiKey = kpiKey;
    this._page = 1;
    this._sortColumn = 'data';
    this._sortDirection = 'asc';

    const includeAll = (kpiKey === 'quantidadeLancamentos' || kpiKey === 'ticketMedio');
    const data = (typeof DataLayer !== 'undefined') ? DataLayer.getData(includeAll) : [];
    this._records = data.filter(config.filter);

    const titleEl = document.getElementById('kpi-detail-title');
    if (titleEl) titleEl.textContent = config.title;

    this._renderSummary();
    this._renderTable();
  },

  close() {
    const modal = document.getElementById('kpi-detail-modal');
    if (modal) {
      modal.setAttribute('hidden', '');
      modal.removeAttribute('style');
    }
    document.body.style.overflow = '';
  },

  _kpiDescriptions: {
    totalRecebido: 'Soma de todas as receitas com status Pago (conciliadas).',
    totalPago: 'Soma de todas as despesas com status Pago (conciliadas).',
    saldo: 'Diferença entre receitas e despesas conciliadas (Pago).',
    recebimentosAberto: 'Receitas com status Pendente que ainda não foram conciliadas.',
    pagamentosAberto: 'Despesas com status Pendente que ainda não foram pagas.',
    fluxoMes: 'Todos os lançamentos do mês atual (receitas - despesas).',
    ticketMedio: 'Valor médio por lançamento considerando todos os registros.',
    quantidadeLancamentos: 'Total de lançamentos financeiros importados.'
  },

  _renderSummary() {
    const summaryEl = document.getElementById('kpi-detail-summary');
    if (!summaryEl) return;

    const desc = this._kpiDescriptions[this._kpiKey] || '';
    const total = this._records.reduce((s, r) => s + r.valor, 0);
    const receitas = this._records.filter(r => r.tipo === 'receita').reduce((s, r) => s + r.valor, 0);
    const despesas = this._records.filter(r => r.tipo === 'despesa').reduce((s, r) => s + r.valor, 0);

    let html = '';
    if (desc) html += `<p style="margin:0 0 8px;color:#64748b;font-size:13px;width:100%;">${desc}</p>`;
    html += `<span><strong>${this._records.length}</strong> lançamento(s)</span>`;
    html += `<span>Total: <strong>${KPICalculator.formatBRL(total)}</strong></span>`;
    if (receitas > 0) html += `<span style="color:var(--success-color,#16a34a)">Receitas: ${KPICalculator.formatBRL(receitas)}</span>`;
    if (despesas > 0) html += `<span style="color:var(--danger-color,#dc2626)">Despesas: ${KPICalculator.formatBRL(despesas)}</span>`;

    summaryEl.innerHTML = html;
  },

  _sortRecords(records, column, direction) {
    const dir = direction === 'desc' ? -1 : 1;
    return [...records].sort((a, b) => {
      const valA = a[column];
      const valB = b[column];

      // Nulls/empties always last
      const aEmpty = (valA == null || valA === '');
      const bEmpty = (valB == null || valB === '');
      if (aEmpty && bEmpty) return 0;
      if (aEmpty) return 1;
      if (bEmpty) return -1;

      // Numeric column
      if (column === 'valor') {
        return (Number(valA) - Number(valB)) * dir;
      }
      // Date column
      if (column === 'data') {
        const dateA = valA instanceof Date ? valA.getTime() : new Date(valA).getTime();
        const dateB = valB instanceof Date ? valB.getTime() : new Date(valB).getTime();
        return (dateA - dateB) * dir;
      }
      // Text columns - locale-aware
      return String(valA).localeCompare(String(valB), 'pt-BR') * dir;
    });
  },

  _columns: [
    { key: 'data', label: 'Data' },
    { key: 'projeto', label: 'Projeto' },
    { key: 'cliente', label: 'Cliente' },
    { key: 'descricao', label: 'Descrição' },
    { key: 'status', label: 'Status' },
    { key: 'valor', label: 'Valor' }
  ],

  _handleHeaderClick(columnKey) {
    if (this._sortColumn === columnKey) {
      this._sortDirection = this._sortDirection === 'asc' ? 'desc' : 'asc';
    } else {
      this._sortColumn = columnKey;
      this._sortDirection = 'asc';
    }
    this._page = 1;
    this._renderTable();
  },

  _renderTableHeader() {
    const thead = document.querySelector('#table-kpi-detail thead tr');
    if (!thead) return;
    thead.innerHTML = '';
    for (const col of this._columns) {
      const th = document.createElement('th');
      th.setAttribute('data-column', col.key);
      th.setAttribute('role', 'columnheader');
      th.style.cursor = 'pointer';
      th.setAttribute('tabindex', '0');

      const label = document.createElement('span');
      label.textContent = col.label;
      th.appendChild(label);

      const indicator = document.createElement('span');
      indicator.className = 'sort-indicator';
      indicator.setAttribute('aria-hidden', 'true');
      if (this._sortColumn === col.key) {
        indicator.textContent = this._sortDirection === 'asc' ? ' ↑' : ' ↓';
        th.setAttribute('aria-sort', this._sortDirection === 'asc' ? 'ascending' : 'descending');
      }
      th.appendChild(indicator);

      th.addEventListener('click', () => this._handleHeaderClick(col.key));
      thead.appendChild(th);
    }
  },

  _renderTable() {
    const tbody = document.querySelector('#table-kpi-detail tbody');
    if (!tbody) return;

    // Update header sort indicators
    this._renderTableHeader();

    // Sort records before pagination
    const sorted = this._sortRecords(this._records, this._sortColumn, this._sortDirection);

    const totalPages = Math.max(1, Math.ceil(sorted.length / this._pageSize));
    if (this._page > totalPages) this._page = totalPages;

    const start = (this._page - 1) * this._pageSize;
    const end = Math.min(start + this._pageSize, sorted.length);
    const pageRecords = sorted.slice(start, end);

    if (pageRecords.length === 0) {
      const emptyMsg = this._kpiKey === 'custom' && this._records.length === 0
        ? 'Nenhum lançamento encontrado para o elemento selecionado.'
        : 'Nenhum lançamento.';
      tbody.innerHTML = `<tr><td colspan="6" style="text-align:center;padding:20px;color:var(--text-muted);">${emptyMsg}</td></tr>`;
    } else {
      let html = '';
      for (const record of pageRecords) {
        const dateStr = record.data instanceof Date ? TableModule.formatDate(record.data) : '';
        const valueFormatted = KPICalculator.formatBRL(record.valor);
        const colorClass = record.tipo === 'receita' ? 'positive' : 'negative';
        const statusBadge = record.status === 'Pago' ? 'badge-pago' : 'badge-pendente';

        html += `<tr data-record-id="${record.id}">`;
        html += `<td>${dateStr}</td>`;
        html += `<td title="${this._esc(record.projeto)}">${this._esc((record.projeto || '-').substring(0, 25))}</td>`;
        html += `<td title="${this._esc(record.cliente)}">${this._esc((record.cliente || '-').substring(0, 25))}</td>`;
        html += `<td title="${this._esc(record.descricao)}">${this._esc((record.descricao || '-').substring(0, 30))}</td>`;
        html += `<td><span class="record-badge ${statusBadge}">${record.status}</span></td>`;
        html += `<td class="${colorClass}">${valueFormatted}</td>`;
        html += `</tr>`;
      }
      tbody.innerHTML = html;
    }

    // Multi-page pagination
    const paginationContainer = document.getElementById('kpi-detail-pagination');
    renderPagination(paginationContainer, this._page, totalPages, (newPage) => {
      this._page = Math.max(1, Math.min(newPage, totalPages));
      this._renderTable();
    });
  },

  _esc(s) { return s ? String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;') : ''; },

  init() {
    // Close button
    const btnClose = document.getElementById('btn-close-kpi-detail');
    if (btnClose) btnClose.addEventListener('click', () => this.close());

    // Overlay click
    const modal = document.getElementById('kpi-detail-modal');
    if (modal) modal.addEventListener('click', (e) => { if (e.target === modal) this.close(); });

    // Pagination is now handled dynamically by renderPagination() in _renderTable()

    // Row click to open record detail
    const tbody = document.querySelector('#table-kpi-detail tbody');
    if (tbody) {
      tbody.addEventListener('click', (e) => {
        const row = e.target.closest('tr[data-record-id]');
        if (!row) return;
        const recordId = row.dataset.recordId;
        const allData = DataLayer.getData();
        const record = allData.find(r => r.id === recordId);
        if (record) RecordDetail.open(record);
      });
    }

    // KPI card click via event delegation on kpi-grid
    const kpiGrid = document.querySelector('.kpi-grid');
    if (kpiGrid) {
      kpiGrid.addEventListener('click', (e) => {
        const card = e.target.closest('.kpi-card[data-kpi]');
        if (!card) return;
        const kpiKey = card.dataset.kpi;
        if (kpiKey) this.open(kpiKey);
      });
    }
  }
};

if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', () => KPIDetailModal.init());
} else {
  KPIDetailModal.init();
}
// Make globally accessible for inline scripts
window.KPIDetailModal = KPIDetailModal;

// ============================================================
// SECTION: Alerts Modal (Full List with Filters)
// ============================================================

const AlertsModal = {
  _page: 1,
  _pageSize: 20,
  _filteredRecords: [],
  _filters: {},

  open() {
    const modal = document.getElementById('alerts-modal');
    if (!modal) return;

    // Populate filter dropdowns
    this._populateFilters();

    // Reset filters and page
    this._filters = {};
    this._page = 1;
    this._clearFilterInputs();

    // Load data
    this._applyFilters();

    modal.hidden = false;
    document.body.style.overflow = 'hidden';
  },

  close() {
    const modal = document.getElementById('alerts-modal');
    if (modal) modal.hidden = true;
    document.body.style.overflow = '';
  },

  _populateFilters() {
    const allData = DataLayer.getData();
    const pendingRecords = allData.filter(r => r.status === 'Pendente');

    // Projetos
    const projetos = [...new Set(pendingRecords.map(r => r.projeto).filter(Boolean))].sort((a, b) => a.localeCompare(b, 'pt-BR'));
    const projSelect = document.getElementById('alerts-filter-projeto');
    if (projSelect) {
      projSelect.innerHTML = '<option value="">Todos</option>' + projetos.map(p => `<option value="${this._esc(p)}">${this._esc(p)}</option>`).join('');
    }

    // Clientes
    const clientes = [...new Set(pendingRecords.map(r => r.cliente).filter(Boolean))].sort((a, b) => a.localeCompare(b, 'pt-BR'));
    const cliSelect = document.getElementById('alerts-filter-cliente');
    if (cliSelect) {
      cliSelect.innerHTML = '<option value="">Todos</option>' + clientes.map(c => `<option value="${this._esc(c)}">${this._esc(c)}</option>`).join('');
    }

    // Contas
    const contas = [...new Set(pendingRecords.map(r => r.conta).filter(Boolean))].sort((a, b) => a.localeCompare(b, 'pt-BR'));
    const contaSelect = document.getElementById('alerts-filter-conta');
    if (contaSelect) {
      contaSelect.innerHTML = '<option value="">Todas</option>' + contas.map(c => `<option value="${this._esc(c)}">${this._esc(c)}</option>`).join('');
    }
  },

  _clearFilterInputs() {
    const ids = ['alerts-filter-periodo', 'alerts-filter-projeto', 'alerts-filter-cliente', 'alerts-filter-tipo', 'alerts-filter-conta'];
    ids.forEach(id => {
      const el = document.getElementById(id);
      if (el) el.value = '';
    });
    // Reset calendar picker state
    this._periodoStart = null;
    this._periodoEnd = null;
  },

  _readFilters() {
    const filters = {};
    const projeto = document.getElementById('alerts-filter-projeto');
    const cliente = document.getElementById('alerts-filter-cliente');
    const tipo = document.getElementById('alerts-filter-tipo');
    const conta = document.getElementById('alerts-filter-conta');

    if (this._periodoStart) filters.dataInicio = this._periodoStart;
    if (this._periodoEnd) filters.dataFim = this._periodoEnd;
    if (projeto && projeto.value) filters.projeto = projeto.value;
    if (cliente && cliente.value) filters.cliente = cliente.value;
    if (tipo && tipo.value) filters.tipo = tipo.value;
    if (conta && conta.value) filters.conta = conta.value;

    return filters;
  },

  _applyFilters() {
    this._filters = this._readFilters();
    const allData = DataLayer.getData();
    let records = allData.filter(r => r.status === 'Pendente');

    // Apply filters
    if (this._filters.dataInicio) {
      records = records.filter(r => {
        const d = r.data instanceof Date ? r.data : new Date(r.data);
        return d >= this._filters.dataInicio;
      });
    }
    if (this._filters.dataFim) {
      records = records.filter(r => {
        const d = r.data instanceof Date ? r.data : new Date(r.data);
        return d <= this._filters.dataFim;
      });
    }
    if (this._filters.projeto) records = records.filter(r => r.projeto === this._filters.projeto);
    if (this._filters.cliente) records = records.filter(r => r.cliente === this._filters.cliente);
    if (this._filters.tipo) records = records.filter(r => r.tipo === this._filters.tipo);
    if (this._filters.conta) records = records.filter(r => r.conta === this._filters.conta);

    // Sort by date (oldest first)
    records.sort((a, b) => {
      const da = a.data instanceof Date ? a.data : new Date(a.data);
      const db = b.data instanceof Date ? b.data : new Date(b.data);
      return da - db;
    });

    this._filteredRecords = records;
    this._page = 1;
    this._renderTable();
  },

  _renderTable() {
    const tbody = document.querySelector('#table-alerts-modal tbody');
    if (!tbody) return;

    const totalPages = Math.max(1, Math.ceil(this._filteredRecords.length / this._pageSize));
    if (this._page > totalPages) this._page = totalPages;

    const start = (this._page - 1) * this._pageSize;
    const end = Math.min(start + this._pageSize, this._filteredRecords.length);
    const pageRecords = this._filteredRecords.slice(start, end);

    // Summary
    const summaryEl = document.getElementById('alerts-modal-summary');
    if (summaryEl) {
      const totalReceita = this._filteredRecords.filter(r => r.tipo === 'receita').reduce((s, r) => s + r.valor, 0);
      const totalDespesa = this._filteredRecords.filter(r => r.tipo === 'despesa').reduce((s, r) => s + r.valor, 0);
      summaryEl.innerHTML = `<strong>${this._filteredRecords.length}</strong> lançamento(s) pendente(s) &nbsp;|&nbsp; ` +
        `<span style="color:var(--success-color,#16a34a)">Receitas: ${KPICalculator.formatBRL(totalReceita)}</span> &nbsp;|&nbsp; ` +
        `<span style="color:var(--danger-color,#dc2626)">Despesas: ${KPICalculator.formatBRL(totalDespesa)}</span>`;
    }

    // Table body
    if (pageRecords.length === 0) {
      tbody.innerHTML = '<tr><td colspan="6" style="text-align:center;padding:20px;color:var(--text-muted);">Nenhum lançamento encontrado.</td></tr>';
    } else {
      let html = '';
      for (const record of pageRecords) {
        const dateStr = record.data instanceof Date ? TableModule.formatDate(record.data) : '';
        const valueFormatted = KPICalculator.formatBRL(record.valor);
        const colorClass = record.tipo === 'receita' ? 'positive' : 'negative';
        const tipoLabel = record.tipo === 'receita' ? 'Receita' : 'Despesa';
        const desc = record.descricao || '-';

        html += `<tr data-record-id="${record.id}">`;
        html += `<td>${dateStr}</td>`;
        html += `<td title="${this._esc(record.projeto)}">${this._esc((record.projeto || '-').substring(0, 30))}</td>`;
        html += `<td title="${this._esc(record.cliente)}">${this._esc((record.cliente || '-').substring(0, 30))}</td>`;
        html += `<td title="${this._esc(desc)}">${this._esc(desc.substring(0, 35))}</td>`;
        html += `<td><span class="record-badge badge-${record.tipo}">${tipoLabel}</span></td>`;
        html += `<td class="${colorClass}">${valueFormatted}</td>`;
        html += `</tr>`;
      }
      tbody.innerHTML = html;
    }

    // Multi-page pagination
    const paginationContainer = document.getElementById('alerts-modal-pagination');
    renderPagination(paginationContainer, this._page, totalPages, (newPage) => {
      this._page = Math.max(1, Math.min(newPage, totalPages));
      this._renderTable();
    });
  },

  _esc(s) { return s ? String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;') : ''; },

  init() {
    // Close button
    const btnClose = document.getElementById('btn-close-alerts-modal');
    if (btnClose) btnClose.addEventListener('click', () => this.close());

    // Close on overlay click
    const modal = document.getElementById('alerts-modal');
    if (modal) modal.addEventListener('click', (e) => { if (e.target === modal) this.close(); });

    // Filter change listeners
    const filterIds = ['alerts-filter-projeto', 'alerts-filter-cliente', 'alerts-filter-tipo', 'alerts-filter-conta'];
    filterIds.forEach(id => {
      const el = document.getElementById(id);
      if (el) el.addEventListener('change', () => this._applyFilters());
    });

    // Periodo picker for alerts
    this._periodoStart = null;
    this._periodoEnd = null;
    this._periodoSelecting = false;
    this._periodoViewYear = new Date().getFullYear();
    this._periodoViewMonth = new Date().getMonth();
    this._periodoMode = 'days'; // 'days', 'months', 'years'

    const periodoInput = document.getElementById('alerts-filter-periodo');
    const periodoPicker = document.getElementById('alerts-periodo-picker');

    if (periodoInput && periodoPicker) {
      periodoInput.addEventListener('click', (e) => { e.stopPropagation(); periodoPicker.hidden = !periodoPicker.hidden; if (!periodoPicker.hidden) this._renderAlertsDays(); });
      periodoPicker.addEventListener('click', (e) => e.stopPropagation());

      document.getElementById('alerts-periodo-prev-month').addEventListener('click', (e) => { e.stopPropagation(); this._periodoViewMonth--; if (this._periodoViewMonth < 0) { this._periodoViewMonth = 11; this._periodoViewYear--; } this._renderAlertsDays(); });
      document.getElementById('alerts-periodo-next-month').addEventListener('click', (e) => { e.stopPropagation(); this._periodoViewMonth++; if (this._periodoViewMonth > 11) { this._periodoViewMonth = 0; this._periodoViewYear++; } this._renderAlertsDays(); });

      // Click month name to show months grid
      document.getElementById('alerts-periodo-month-name').addEventListener('click', (e) => {
        e.stopPropagation();
        this._renderAlertsMonths();
      });

      // Click year name to show years grid
      document.getElementById('alerts-periodo-year-name').addEventListener('click', (e) => {
        e.stopPropagation();
        this._renderAlertsYears();
      });

      document.getElementById('alerts-periodo-clear').addEventListener('click', (e) => { e.stopPropagation(); this._periodoStart = null; this._periodoEnd = null; this._periodoSelecting = false; periodoInput.value = ''; this._applyFilters(); periodoPicker.hidden = true; });
      document.getElementById('alerts-periodo-apply').addEventListener('click', (e) => { e.stopPropagation(); this._updateAlertsPeriodoInput(); this._applyFilters(); periodoPicker.hidden = true; });

      // Close on outside click
      document.addEventListener('click', () => { periodoPicker.hidden = true; });
    }

    // Clear filters button
    const btnClear = document.getElementById('btn-alerts-clear-filters');
    if (btnClear) btnClear.addEventListener('click', () => {
      this._clearFilterInputs();
      this._applyFilters();
    });

    // Pagination is now handled by renderPagination() in _renderTable()

    // Click row to open record detail
    const tbody = document.querySelector('#table-alerts-modal tbody');
    if (tbody) {
      tbody.addEventListener('click', (e) => {
        const row = e.target.closest('tr[data-record-id]');
        if (!row) return;
        const recordId = row.dataset.recordId;
        const allData = DataLayer.getData();
        const record = allData.find(r => r.id === recordId);
        if (record) RecordDetail.open(record);
      });
    }
  },

  // --- Alerts Calendar Picker Methods ---
  _renderAlertsDays() {
    this._periodoMode = 'days';
    const monthNameEl = document.getElementById('alerts-periodo-month-name');
    const yearNameEl = document.getElementById('alerts-periodo-year-name');
    const container = document.getElementById('alerts-periodo-days');
    const weekdays = document.getElementById('alerts-periodo-weekdays');
    const yearGrid = document.getElementById('alerts-periodo-year-grid');
    const monthGrid = document.getElementById('alerts-periodo-month-grid');
    const monthNames = ['Janeiro','Fevereiro','Março','Abril','Maio','Junho','Julho','Agosto','Setembro','Outubro','Novembro','Dezembro'];

    if (yearGrid) yearGrid.hidden = true;
    if (monthGrid) monthGrid.hidden = true;
    if (weekdays) weekdays.hidden = false;
    if (container) container.hidden = false;
    if (monthNameEl) monthNameEl.textContent = monthNames[this._periodoViewMonth];
    if (yearNameEl) yearNameEl.textContent = this._periodoViewYear;

    const firstDay = new Date(this._periodoViewYear, this._periodoViewMonth, 1).getDay();
    const daysInMonth = new Date(this._periodoViewYear, this._periodoViewMonth + 1, 0).getDate();
    const today = new Date(); today.setHours(0,0,0,0);

    let html = '';
    for (let i = 0; i < firstDay; i++) html += '<button class="periodo-day empty" disabled></button>';
    for (let d = 1; d <= daysInMonth; d++) {
      const date = new Date(this._periodoViewYear, this._periodoViewMonth, d); date.setHours(0,0,0,0);
      let cls = 'periodo-day';
      if (date.getTime() === today.getTime()) cls += ' today';
      if (this._periodoStart && date.getTime() === this._periodoStart.getTime()) cls += ' selected';
      if (this._periodoEnd && date.getTime() === this._periodoEnd.getTime()) cls += ' selected';
      if (this._periodoStart && this._periodoEnd && date > this._periodoStart && date < this._periodoEnd) cls += ' in-range';
      html += `<button class="${cls}" data-date="${this._periodoViewYear}-${String(this._periodoViewMonth+1).padStart(2,'0')}-${String(d).padStart(2,'0')}">${d}</button>`;
    }
    if (container) container.innerHTML = html;

    // Day click
    if (container) container.querySelectorAll('.periodo-day:not(.empty)').forEach(btn => {
      btn.addEventListener('click', () => {
        const clicked = new Date(btn.dataset.date + 'T00:00:00');
        if (!this._periodoStart || !this._periodoSelecting) {
          this._periodoStart = clicked; this._periodoEnd = null; this._periodoSelecting = true;
        } else {
          if (clicked < this._periodoStart) { this._periodoEnd = this._periodoStart; this._periodoStart = clicked; }
          else { this._periodoEnd = clicked; }
          this._periodoSelecting = false;
        }
        this._renderAlertsDays();
      });
    });
  },

  _renderAlertsYears() {
    this._periodoMode = 'years';
    const container = document.getElementById('alerts-periodo-days');
    const weekdays = document.getElementById('alerts-periodo-weekdays');
    const yearGrid = document.getElementById('alerts-periodo-year-grid');
    const monthGrid = document.getElementById('alerts-periodo-month-grid');

    if (container) container.hidden = true;
    if (weekdays) weekdays.hidden = true;
    if (monthGrid) monthGrid.hidden = true;
    if (yearGrid) yearGrid.hidden = false;

    const currentYear = new Date().getFullYear();
    let html = '';
    for (let y = currentYear - 5; y <= currentYear + 5; y++) {
      const cls = y === this._periodoViewYear ? 'periodo-year-btn active' : 'periodo-year-btn';
      html += `<button class="${cls}" data-year="${y}">${y}</button>`;
    }
    if (yearGrid) yearGrid.innerHTML = html;

    yearGrid.querySelectorAll('.periodo-year-btn').forEach(btn => {
      btn.addEventListener('click', (e) => {
        e.stopPropagation();
        this._periodoViewYear = parseInt(btn.dataset.year);
        this._renderAlertsDays();
      });
    });
  },

  _renderAlertsMonths() {
    this._periodoMode = 'months';
    const container = document.getElementById('alerts-periodo-days');
    const weekdays = document.getElementById('alerts-periodo-weekdays');
    const yearGrid = document.getElementById('alerts-periodo-year-grid');
    const monthGrid = document.getElementById('alerts-periodo-month-grid');

    if (container) container.hidden = true;
    if (weekdays) weekdays.hidden = true;
    if (yearGrid) yearGrid.hidden = true;
    if (monthGrid) monthGrid.hidden = false;

    const monthNames = ['Jan','Fev','Mar','Abr','Mai','Jun','Jul','Ago','Set','Out','Nov','Dez'];
    let html = '';
    for (let m = 0; m < 12; m++) {
      const cls = m === this._periodoViewMonth ? 'periodo-year-btn active' : 'periodo-year-btn';
      html += `<button class="${cls}" data-month="${m}">${monthNames[m]}</button>`;
    }
    if (monthGrid) monthGrid.innerHTML = html;

    monthGrid.querySelectorAll('.periodo-year-btn').forEach(btn => {
      btn.addEventListener('click', (e) => {
        e.stopPropagation();
        this._periodoViewMonth = parseInt(btn.dataset.month);
        this._renderAlertsDays();
      });
    });
  },

  _updateAlertsPeriodoInput() {
    const input = document.getElementById('alerts-filter-periodo');
    if (!input) return;
    if (!this._periodoStart) { input.value = ''; return; }
    const fmt = (d) => `${String(d.getDate()).padStart(2,'0')}/${String(d.getMonth()+1).padStart(2,'0')}/${d.getFullYear()}`;
    if (this._periodoEnd && this._periodoStart.getTime() !== this._periodoEnd.getTime()) {
      input.value = `${fmt(this._periodoStart)} - ${fmt(this._periodoEnd)}`;
    } else {
      input.value = fmt(this._periodoStart);
    }
  }
};

// Initialize AlertsModal on DOM ready
document.addEventListener('DOMContentLoaded', () => AlertsModal.init());

// ============================================================
// SECTION: Projeto Detalhe (Project Detail View)
// ============================================================

const ProjetoDetalhe = {
  _selectedProjeto: '',
  _statusFilter: '',
  _records: [],
  _page: 1,
  _pageSize: 25,
  _activeTab: 'projeto-previsto-realizado',

  init() {
    const selectProjeto = document.getElementById('projeto-detalhe-select');
    const selectStatus = document.getElementById('projeto-detalhe-status');

    if (selectProjeto) {
      selectProjeto.addEventListener('change', () => {
        this._selectedProjeto = selectProjeto.value;
        this._page = 1;
        this._refresh();
      });
    }

    if (selectStatus) {
      selectStatus.addEventListener('change', () => {
        this._statusFilter = selectStatus.value;
        this._page = 1;
        this._refresh();
      });
    }

    // Tab switching
    const tabs = document.querySelectorAll('#projeto-detalhe-tabs .detail-tab');
    tabs.forEach(tab => {
      tab.addEventListener('click', () => {
        tabs.forEach(t => t.classList.remove('active'));
        tab.classList.add('active');
        this._activeTab = tab.dataset.tab;
        this._showActiveTab();
      });
    });

    // Pagination for lancamentos tab
    const paginationEl = document.getElementById('projeto-lancamentos-pagination');
    if (paginationEl) {
      const prevBtn = paginationEl.querySelector('.btn-page-prev');
      const nextBtn = paginationEl.querySelector('.btn-page-next');
      if (prevBtn) prevBtn.addEventListener('click', () => { if (this._page > 1) { this._page--; this._renderLancamentos(); } });
      if (nextBtn) nextBtn.addEventListener('click', () => {
        const totalPages = Math.ceil(this._records.length / this._pageSize);
        if (this._page < totalPages) { this._page++; this._renderLancamentos(); }
      });
    }

    // Click row to open record detail
    const tbody = document.querySelector('#table-projeto-lancamentos tbody');
    if (tbody) {
      tbody.addEventListener('click', (e) => {
        const row = e.target.closest('tr[data-record-id]');
        if (!row) return;
        const recordId = row.dataset.recordId;
        const allData = DataLayer.getData();
        const record = allData.find(r => r.id === recordId);
        if (record) RecordDetail.open(record);
      });
    }
  },

  populateProjects() {
    const allData = DataLayer.getData();
    const projetos = [...new Set(allData.map(r => r.projeto).filter(Boolean))].sort((a, b) => a.localeCompare(b, 'pt-BR'));
    const select = document.getElementById('projeto-detalhe-select');
    if (select) {
      select.innerHTML = '<option value="">-- Selecione um projeto --</option>' +
        projetos.map(p => `<option value="${this._esc(p)}">${this._esc(p)}</option>`).join('');
    }
  },

  _refresh() {
    const emptyEl = document.getElementById('projeto-detalhe-empty');
    const kpisEl = document.getElementById('projeto-detalhe-kpis');
    const tabsEl = document.getElementById('projeto-detalhe-tabs');
    const tabPR = document.getElementById('tab-projeto-previsto-realizado');
    const tabLanc = document.getElementById('tab-projeto-lancamentos');

    if (!this._selectedProjeto) {
      // Show empty state
      if (emptyEl) emptyEl.hidden = false;
      if (kpisEl) kpisEl.hidden = true;
      if (tabsEl) tabsEl.hidden = true;
      if (tabPR) tabPR.hidden = true;
      if (tabLanc) tabLanc.hidden = true;
      return;
    }

    // Hide empty, show content
    if (emptyEl) emptyEl.hidden = true;
    if (kpisEl) kpisEl.hidden = false;
    if (tabsEl) tabsEl.hidden = false;

    // Get records for this project
    const allData = DataLayer.getData();
    let records = allData.filter(r => r.projeto === this._selectedProjeto);

    // Apply status filter
    if (this._statusFilter) {
      records = records.filter(r => r.status === this._statusFilter);
    }

    this._records = records;

    // Calculate KPIs
    this._renderKPIs(allData.filter(r => r.projeto === this._selectedProjeto));

    // Show active tab
    this._showActiveTab();
  },

  _renderKPIs(allProjectRecords) {
    // Previsto = sum of ALL records (Pendente + Pago)
    // Realizado = sum of only Pago (conciliado) records
    const pagos = allProjectRecords.filter(r => r.status === 'Pago');

    const previsto = allProjectRecords.reduce((s, r) => s + r.valor, 0);
    const realizado = pagos.reduce((s, r) => s + r.valor, 0);
    const saldo = realizado - previsto;
    const pendente = allProjectRecords.filter(r => r.status === 'Pendente').reduce((s, r) => s + r.valor, 0);

    const elPrevisto = document.getElementById('projeto-kpi-previsto');
    const elRealizado = document.getElementById('projeto-kpi-realizado');
    const elSaldo = document.getElementById('projeto-kpi-saldo');
    const elPendente = document.getElementById('projeto-kpi-pendente');

    if (elPrevisto) elPrevisto.textContent = KPICalculator.formatBRL(previsto);
    if (elRealizado) elRealizado.textContent = KPICalculator.formatBRL(realizado);
    if (elSaldo) elSaldo.textContent = KPICalculator.formatBRL(saldo);
    if (elPendente) elPendente.textContent = KPICalculator.formatBRL(pendente);
  },

  _showActiveTab() {
    const tabPR = document.getElementById('tab-projeto-previsto-realizado');
    const tabLanc = document.getElementById('tab-projeto-lancamentos');

    if (this._activeTab === 'projeto-previsto-realizado') {
      if (tabPR) tabPR.hidden = false;
      if (tabLanc) tabLanc.hidden = true;
      this._renderPrevistoRealizado();
    } else {
      if (tabPR) tabPR.hidden = true;
      if (tabLanc) tabLanc.hidden = false;
      this._renderLancamentos();
    }
  },

  _renderPrevistoRealizado() {
    const container = document.getElementById('projeto-pr-tree');
    if (!container) return;

    // Get ALL records for this project (ignore status filter for this view)
    const allData = DataLayer.getData();
    const projectRecords = allData.filter(r => r.projeto === this._selectedProjeto);

    if (projectRecords.length === 0) {
      container.innerHTML = '<p style="text-align:center;padding:20px;color:var(--text-muted);">Nenhum lançamento encontrado.</p>';
      return;
    }

    // Build hierarchy: tipoClassif (1ª Cat) → nivel1 (2ª Cat) → nivel2 (3ª Cat)
    const tree = {};
    for (const record of projectRecords) {
      const lvl1 = record.tipoClassif || 'Sem Classificação';
      const lvl2 = record.nivel1 || 'Sem 2ª Categoria';
      const lvl3 = record.nivel2 || 'Sem 3ª Categoria';

      if (!tree[lvl1]) tree[lvl1] = { previsto: 0, realizado: 0, children: {} };
      if (!tree[lvl1].children[lvl2]) tree[lvl1].children[lvl2] = { previsto: 0, realizado: 0, children: {} };
      if (!tree[lvl1].children[lvl2].children[lvl3]) tree[lvl1].children[lvl2].children[lvl3] = { previsto: 0, realizado: 0 };

      const valor = record.valor;
      // Previsto = ALL records
      tree[lvl1].previsto += valor;
      tree[lvl1].children[lvl2].previsto += valor;
      tree[lvl1].children[lvl2].children[lvl3].previsto += valor;
      // Realizado = only Pago
      if (record.status === 'Pago') {
        tree[lvl1].realizado += valor;
        tree[lvl1].children[lvl2].realizado += valor;
        tree[lvl1].children[lvl2].children[lvl3].realizado += valor;
      }
    }

    // Render tree
    let totalPrevisto = 0, totalRealizado = 0;
    let html = '<div class="pr-tree-header"><span class="pr-tree-col-name">Categoria</span><span class="pr-tree-col">Previsto</span><span class="pr-tree-col">Realizado</span><span class="pr-tree-col">Diferença</span></div>';

    const sortedLvl1 = Object.entries(tree).sort((a, b) => a[0].localeCompare(b[0], 'pt-BR'));

    for (const [lvl1Name, lvl1Data] of sortedLvl1) {
      totalPrevisto += lvl1Data.previsto;
      totalRealizado += lvl1Data.realizado;
      const diff1 = lvl1Data.realizado - lvl1Data.previsto;
      const diffClass1 = diff1 >= 0 ? 'positive' : 'negative';

      html += `<div class="pr-tree-node pr-tree-lvl1" data-expanded="false">`;
      html += `<div class="pr-tree-row pr-tree-row-clickable" data-level="1">`;
      html += `<span class="pr-tree-col-name"><span class="pr-tree-toggle">▶</span> ${this._esc(lvl1Name)}</span>`;
      html += `<span class="pr-tree-col">${KPICalculator.formatBRL(lvl1Data.previsto)}</span>`;
      html += `<span class="pr-tree-col">${KPICalculator.formatBRL(lvl1Data.realizado)}</span>`;
      html += `<span class="pr-tree-col ${diffClass1}">${KPICalculator.formatBRL(diff1)}</span>`;
      html += `</div>`;

      // Level 2 children
      html += `<div class="pr-tree-children" hidden>`;
      const sortedLvl2 = Object.entries(lvl1Data.children).sort((a, b) => a[0].localeCompare(b[0], 'pt-BR'));

      for (const [lvl2Name, lvl2Data] of sortedLvl2) {
        const diff2 = lvl2Data.realizado - lvl2Data.previsto;
        const diffClass2 = diff2 >= 0 ? 'positive' : 'negative';

        html += `<div class="pr-tree-node pr-tree-lvl2" data-expanded="false">`;
        html += `<div class="pr-tree-row pr-tree-row-clickable" data-level="2">`;
        html += `<span class="pr-tree-col-name"><span class="pr-tree-toggle">▶</span> ${this._esc(lvl2Name)}</span>`;
        html += `<span class="pr-tree-col">${KPICalculator.formatBRL(lvl2Data.previsto)}</span>`;
        html += `<span class="pr-tree-col">${KPICalculator.formatBRL(lvl2Data.realizado)}</span>`;
        html += `<span class="pr-tree-col ${diffClass2}">${KPICalculator.formatBRL(diff2)}</span>`;
        html += `</div>`;

        // Level 3 children
        html += `<div class="pr-tree-children" hidden>`;
        const sortedLvl3 = Object.entries(lvl2Data.children).sort((a, b) => a[0].localeCompare(b[0], 'pt-BR'));

        for (const [lvl3Name, lvl3Data] of sortedLvl3) {
          const diff3 = lvl3Data.realizado - lvl3Data.previsto;
          const diffClass3 = diff3 >= 0 ? 'positive' : 'negative';

          html += `<div class="pr-tree-row pr-tree-lvl3">`;
          html += `<span class="pr-tree-col-name">${this._esc(lvl3Name)}</span>`;
          html += `<span class="pr-tree-col">${KPICalculator.formatBRL(lvl3Data.previsto)}</span>`;
          html += `<span class="pr-tree-col">${KPICalculator.formatBRL(lvl3Data.realizado)}</span>`;
          html += `<span class="pr-tree-col ${diffClass3}">${KPICalculator.formatBRL(diff3)}</span>`;
          html += `</div>`;
        }

        html += `</div>`; // close lvl3 children
        html += `</div>`; // close lvl2 node
      }

      html += `</div>`; // close lvl2 children
      html += `</div>`; // close lvl1 node
    }

    // Total row
    const totalDiff = totalRealizado - totalPrevisto;
    const totalDiffClass = totalDiff >= 0 ? 'positive' : 'negative';
    html += `<div class="pr-tree-row pr-tree-total">`;
    html += `<span class="pr-tree-col-name"><strong>TOTAL</strong></span>`;
    html += `<span class="pr-tree-col"><strong>${KPICalculator.formatBRL(totalPrevisto)}</strong></span>`;
    html += `<span class="pr-tree-col"><strong>${KPICalculator.formatBRL(totalRealizado)}</strong></span>`;
    html += `<span class="pr-tree-col ${totalDiffClass}"><strong>${KPICalculator.formatBRL(totalDiff)}</strong></span>`;
    html += `</div>`;

    container.innerHTML = html;

    // Attach toggle click handlers
    container.querySelectorAll('.pr-tree-row-clickable').forEach(row => {
      row.addEventListener('click', () => {
        const node = row.closest('.pr-tree-node');
        if (!node) return;
        const children = node.querySelector(':scope > .pr-tree-children');
        if (!children) return;
        const expanded = node.dataset.expanded === 'true';
        node.dataset.expanded = expanded ? 'false' : 'true';
        children.hidden = expanded;
        const toggle = row.querySelector('.pr-tree-toggle');
        if (toggle) toggle.textContent = expanded ? '▶' : '▼';
      });
    });
  },

  _renderLancamentos() {
    const tbody = document.querySelector('#table-projeto-lancamentos tbody');
    if (!tbody) return;

    const totalPages = Math.max(1, Math.ceil(this._records.length / this._pageSize));
    if (this._page > totalPages) this._page = totalPages;

    const start = (this._page - 1) * this._pageSize;
    const end = Math.min(start + this._pageSize, this._records.length);

    // Sort by date desc
    const sorted = [...this._records].sort((a, b) => {
      const da = a.data instanceof Date ? a.data : new Date(a.data);
      const db = b.data instanceof Date ? b.data : new Date(b.data);
      return db - da;
    });

    const pageRecords = sorted.slice(start, end);

    if (pageRecords.length === 0) {
      tbody.innerHTML = '<tr><td colspan="6" style="text-align:center;padding:20px;color:var(--text-muted);">Nenhum lançamento encontrado.</td></tr>';
    } else {
      let html = '';
      for (const record of pageRecords) {
        const dateStr = record.data instanceof Date ? TableModule.formatDate(record.data) : '';
        const valueFormatted = KPICalculator.formatBRL(record.valor);
        const colorClass = record.tipo === 'receita' ? 'positive' : 'negative';
        const tipoLabel = record.tipo === 'receita' ? 'Receita' : 'Despesa';
        const statusBadge = record.status === 'Pago' ? 'badge-pago' : 'badge-pendente';

        html += `<tr data-record-id="${record.id}">`;
        html += `<td>${dateStr}</td>`;
        html += `<td title="${this._esc(record.descricao)}">${this._esc((record.descricao || '-').substring(0, 40))}</td>`;
        html += `<td>${this._esc(record.nivel1 || record.tipoClassif || '-')}</td>`;
        html += `<td><span class="record-badge ${statusBadge}">${record.status}</span></td>`;
        html += `<td><span class="record-badge badge-${record.tipo}">${tipoLabel}</span></td>`;
        html += `<td class="${colorClass}">${valueFormatted}</td>`;
        html += `</tr>`;
      }
      tbody.innerHTML = html;
    }

    // Pagination
    const paginationEl = document.getElementById('projeto-lancamentos-pagination');
    if (paginationEl) {
      const prevBtn = paginationEl.querySelector('.btn-page-prev');
      const nextBtn = paginationEl.querySelector('.btn-page-next');
      const pageInfo = paginationEl.querySelector('.page-info');
      if (prevBtn) prevBtn.disabled = this._page <= 1;
      if (nextBtn) nextBtn.disabled = this._page >= totalPages;
      if (pageInfo) pageInfo.textContent = `Página ${this._page} de ${totalPages}`;
    }
  },

  _esc(s) { return s ? String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;') : ''; }
};

// Initialize ProjetoDetalhe on DOM ready
document.addEventListener('DOMContentLoaded', () => ProjetoDetalhe.init());

// ============================================================
// SECTION: Plano de Contas Report
// ============================================================

const ReportModule = {
  generateReport(data) {
    if (!data || data.length === 0) return;

    const container = document.getElementById('report-tree');
    if (!container) return;

    const hierarchy = this._buildHierarchy(data);
    let totalCredit = 0;
    let totalDebit = 0;
    let html = '';

    for (const [tipo, tipoData] of hierarchy) {
      totalCredit += tipoData.credit;
      totalDebit += tipoData.debit;

      html += `<div class="report-item report-item--l0">`;
      html += this._header(tipo, tipoData.credit, tipoData.debit);
      html += `<div class="report-item-children">`;

      for (const [cat1, cat1Data] of tipoData.children) {
        html += `<div class="report-item report-item--l1">`;
        html += this._header(cat1, cat1Data.credit, cat1Data.debit);
        html += `<div class="report-item-children">`;

        for (const [cat2, cat2Data] of cat1Data.children) {
          html += `<div class="report-item report-item--l2">`;
          html += this._header(cat2, cat2Data.credit, cat2Data.debit);
          html += `<div class="report-item-children"><ul class="report-transactions">`;

          for (const rec of cat2Data.records) {
            const date = rec.data instanceof Date ? TableModule.formatDate(rec.data) : '';
            const desc = rec.descricao || rec.cliente || '';
            const creditVal = rec.tipo === 'receita' ? KPICalculator.formatBRL(rec.valor) : '-';
            const debitVal = rec.tipo === 'despesa' ? KPICalculator.formatBRL(rec.valor) : '-';
            html += `<li class="report-tx" data-record-id="${rec.id}" style="cursor:pointer;"><span class="report-tx-date">${date}</span><span class="report-tx-desc">${this._esc(desc)}</span><span class="report-tx-credit">${creditVal}</span><span class="report-tx-debit">${debitVal}</span></li>`;
          }

          html += `</ul></div></div>`;
        }
        html += `</div></div>`;
      }
      html += `</div></div>`;
    }

    container.innerHTML = html;

    // Update totals
    const el = (id, v) => { const e = document.getElementById(id); if (e) e.textContent = KPICalculator.formatBRL(v); };
    el('report-total-credit', totalCredit);
    el('report-total-debit', totalDebit);
    el('report-total-net', totalCredit - totalDebit);

    // Toggle click
    container.querySelectorAll('.report-item-header').forEach(header => {
      header.addEventListener('click', () => {
        header.parentElement.classList.toggle('open');
      });
    });

    // Click on report transaction to open record detail
    container.querySelectorAll('.report-tx[data-record-id]').forEach(li => {
      li.addEventListener('click', () => {
        const recordId = li.dataset.recordId;
        const allData = DataLayer.getData(true);
        const record = allData.find(r => r.id === recordId);
        if (record) RecordDetail.open(record);
      });
    });

    this._populateContaFilter(data);
  },

  _header(name, credit, debit) {
    const c = credit > 0 ? KPICalculator.formatBRL(credit) : '-';
    const d = debit > 0 ? KPICalculator.formatBRL(debit) : '-';
    return `<div class="report-item-header">
      <div class="report-item-left"><span class="report-item-arrow">&#9654;</span><span class="report-item-name">${this._esc(name)}</span></div>
      <div class="report-item-right"><span class="report-item-credit">${c}</span><span class="report-item-debit">${d}</span></div>
    </div>`;
  },

  _buildHierarchy(data) {
    const hierarchy = new Map();
    const sorted = [...data].sort((a, b) => (a.data instanceof Date ? a.data : new Date(a.data)) - (b.data instanceof Date ? b.data : new Date(b.data)));

    for (const record of sorted) {
      const tipo = record.tipoClassif || '(Sem Classificação)';
      const cat1 = record.nivel1 || '(Sem Categoria)';
      const cat2 = record.nivel2 || '(Sem Categoria)';

      if (!hierarchy.has(tipo)) hierarchy.set(tipo, { credit: 0, debit: 0, children: new Map() });
      const tipoNode = hierarchy.get(tipo);
      if (!tipoNode.children.has(cat1)) tipoNode.children.set(cat1, { credit: 0, debit: 0, children: new Map() });
      const cat1Node = tipoNode.children.get(cat1);
      if (!cat1Node.children.has(cat2)) cat1Node.children.set(cat2, { credit: 0, debit: 0, records: [] });
      const cat2Node = cat1Node.children.get(cat2);

      cat2Node.records.push(record);
      const credit = record.tipo === 'receita' ? record.valor : 0;
      const debit = record.tipo === 'despesa' ? record.valor : 0;
      cat2Node.credit += credit; cat2Node.debit += debit;
      cat1Node.credit += credit; cat1Node.debit += debit;
      tipoNode.credit += credit; tipoNode.debit += debit;
    }
    return hierarchy;
  },

  _populateContaFilter(data) {
    const select = document.getElementById('report-conta-filter');
    if (!select) return;

    // Only populate options from full dataset (not filtered)
    const allData = DataLayer.getData();
    const contas = [...new Set(allData.map(r => r.conta).filter(c => c))].sort();
    const currentValue = select.value;

    select.innerHTML = '<option value="all">Todas as Contas</option>';
    contas.forEach(c => { select.innerHTML += `<option value="${this._esc(c)}">${this._esc(c)}</option>`; });

    // Restore selection
    if (currentValue && currentValue !== 'all') {
      select.value = currentValue;
    }

    // Set handler only once
    if (!select._hasHandler) {
      select._hasHandler = true;
      select.addEventListener('change', () => {
        const v = select.value;
        const source = v === 'all' ? DataLayer.getData() : DataLayer.getData().filter(r => r.conta === v);
        this.generateReport(source);
      });
    }
  },

  _esc(s) { return s ? String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;') : ''; }
};

// Wire print button
document.addEventListener('DOMContentLoaded', () => {
  const btnPrint = document.getElementById('btn-print-report');
  if (btnPrint) {
    btnPrint.addEventListener('click', () => {
      window.print();
    });
  }
});

// When data is loaded, populate filter dropdowns and use lazy rendering strategy
// Priority: KPIs + first chart section rendered immediately (within 3s)
// Deferred: remaining chart sections, table sections, insights section
EventBus.on('data:loaded', (payload) => {
  UIController.hideError(); // Clear any previous error on successful load

  // Store data reference for deferred rendering
  UIController._currentData = payload.records;

  // Reset rendered sections tracking for fresh data load
  UIController._sectionRendered.clear();

  // Priority: Populate filters immediately
  UIController.populateFilters();

  // Priority: Render KPI cards immediately (section-dashboard)
  UIController.updateKPICards(payload.records);
  UIController._sectionRendered.add('section-dashboard');

  // Priority: Render first chart section immediately (section-financeiro)
  // Note: Charts are initialized by _processFile after _showDashboard makes sections visible
  // This handles the case where data:loaded is triggered from other sources
  if (!document.getElementById('section-financeiro')?.hidden) {
    ChartEngine.initCharts(payload.records);
  }
  UIController._sectionRendered.add('section-financeiro');

  // Set up lazy rendering for deferred sections (tables, insights)
  UIController.initLazyRendering();
});

// When filters change, check for no results and update rendered sections only
EventBus.on('filters:changed', (payload) => {
  UIController.checkNoResults(payload.filteredData);
  UIController._currentData = payload.filteredData;
  UIController.updateKPICards(payload.filteredData);

  // Only update sections that have already been rendered
  if (UIController._sectionRendered.has('section-projetos')) {
    UIController.refreshTables(payload.filteredData);
  }
  if (UIController._sectionRendered.has('section-insights')) {
    UIController.renderInsights(payload.filteredData);
  }
  if (UIController._sectionRendered.has('section-relatorios')) {
    ReportModule.generateReport(payload.filteredData);
  }
});

// When filters are cleared, hide no-results message and recalculate rendered sections
EventBus.on('filters:cleared', () => {
  const noResults = document.getElementById('filter-no-results');
  if (noResults) noResults.hidden = true;
  const allData = DataLayer.getData();
  UIController._currentData = allData;
  UIController.updateKPICards(allData);

  // Update charts with full data
  if (UIController._sectionRendered.has('section-financeiro')) {
    ChartEngine.updateCharts(allData);
  }

  // Only update sections that have already been rendered
  if (UIController._sectionRendered.has('section-projetos')) {
    UIController.refreshTables(allData);
  }
  if (UIController._sectionRendered.has('section-insights')) {
    UIController.renderInsights(allData);
  }
  if (UIController._sectionRendered.has('section-relatorios')) {
    ReportModule.generateReport(allData);
  }
});

// Wire 'error:occurred' event to UIController.showError
EventBus.on('error:occurred', (payload) => {
  if (payload.type === 'export') {
    // Create a temporary error toast with retry button for export errors
    const existingToast = document.getElementById('export-error-toast');
    if (existingToast) existingToast.remove();

    const toast = document.createElement('div');
    toast.id = 'export-error-toast';
    toast.setAttribute('role', 'alert');
    toast.setAttribute('aria-live', 'assertive');
    toast.style.cssText = 'position:fixed;bottom:20px;right:20px;background:#dc3545;color:#fff;padding:16px 20px;border-radius:8px;box-shadow:0 4px 12px rgba(0,0,0,0.3);z-index:10000;display:flex;align-items:center;gap:12px;font-size:14px;max-width:400px;';

    const msgSpan = document.createElement('span');
    msgSpan.textContent = payload.message || 'Erro ao exportar.';

    const retryBtn = document.createElement('button');
    retryBtn.textContent = 'Tentar novamente';
    retryBtn.setAttribute('aria-label', 'Tentar exportar PDF novamente');
    retryBtn.style.cssText = 'background:#fff;color:#dc3545;border:none;padding:6px 12px;border-radius:4px;cursor:pointer;font-weight:bold;white-space:nowrap;';
    retryBtn.addEventListener('click', () => {
      toast.remove();
      ExportModule.generatePDF();
    });

    const closeBtn = document.createElement('button');
    closeBtn.textContent = '\u2715';
    closeBtn.setAttribute('aria-label', 'Fechar mensagem de erro');
    closeBtn.style.cssText = 'background:none;border:none;color:#fff;font-size:18px;cursor:pointer;padding:0 4px;';
    closeBtn.addEventListener('click', () => toast.remove());

    toast.appendChild(msgSpan);
    toast.appendChild(retryBtn);
    toast.appendChild(closeBtn);
    document.body.appendChild(toast);

    // Auto-dismiss after 10 seconds
    setTimeout(() => {
      if (toast.parentNode) toast.remove();
    }, 10000);
  } else if (payload.type === 'no-columns') {
    // Show expected columns message for unmappable files
    UIController.showExpectedColumnsMessage();
  } else {
    // General error: show in upload-error area with stage and action
    const message = payload.message || UIController._errorMessages[payload.type] || 'Ocorreu um erro inesperado.';
    UIController.showError(message, payload.stage, payload.action);
  }
});

// Wire 'loading:start' event to UIController.showLoadingOverlay
EventBus.on('loading:start', (payload) => {
  UIController.showLoadingOverlay(payload.stage, payload.progress || 0);
});

// Wire 'loading:end' event to UIController.hideLoadingOverlay
EventBus.on('loading:end', () => {
  UIController.hideLoadingOverlay();
});

// Initialize filter bar, sidebar, upload area, and check browser compatibility when DOM is ready
if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', () => {
    UIController.checkBrowserCompatibility();
    UIController.initFilterBar();
    UIController.initSidebar();
    UIController.initUpload();
  });
} else {
  UIController.checkBrowserCompatibility();
  UIController.initFilterBar();
  UIController.initSidebar();
  UIController.initUpload();
}

// ============================================================
// SECTION: Pie Chart Data Preparation
// ============================================================

/**
 * Aggregates filtered financial records by nivel2 field for pie chart rendering.
 * Returns { labels, data, backgroundColor } or null when no valid data exists.
 * - Filters out zero values
 * - Sorts descending by total, caps at 10 categories
 * - Aggregates remaining categories into "Outros"
 * - Truncates labels to 40 characters with "..." suffix
 * Validates: Requirements 3.2, 3.3, 3.5, 3.7
 * @param {Array} filtered - Array of financial record objects with nivel2 and valor fields
 * @returns {{ labels: string[], data: number[], backgroundColor: string[] } | null}
 */
function _preparePieChartData(filtered) {
  const totals = {};
  for (const r of filtered) {
    if (r.nivel2 && r.nivel2.trim() !== '') {
      totals[r.nivel2] = (totals[r.nivel2] || 0) + Math.abs(r.valor);
    }
  }

  // Filter out zero values
  const nonZero = Object.entries(totals).filter(([, v]) => v > 0);
  if (nonZero.length === 0) return null; // signals empty state

  // Sort descending, take top 10
  const sorted = nonZero.sort((a, b) => b[1] - a[1]);
  let chartEntries = sorted.slice(0, 10);

  // Aggregate remaining as "Outros"
  if (sorted.length > 10) {
    const outrosTotal = sorted.slice(10).reduce((sum, [, v]) => sum + v, 0);
    chartEntries.push(['Outros', outrosTotal]);
  }

  // Truncate labels for legend (max 40 chars)
  const labels = chartEntries.map(([name]) =>
    name.length > 40 ? name.substring(0, 37) + '...' : name
  );

  const colors = ['#4f46e5','#06b6d4','#8b5cf6','#f59e0b','#10b981',
                  '#ef4444','#ec4899','#6366f1','#14b8a6','#f97316','#94a3b8'];

  return {
    labels,
    data: chartEntries.map(([, v]) => v),
    backgroundColor: colors.slice(0, chartEntries.length)
  };
}

// ============================================================
// SECTION: Chart Export Data Labels
// ============================================================

/**
 * Calculate percentage for a segment value relative to total.
 * @param {number} value - Segment value
 * @param {number} total - Sum of all segments
 * @returns {string} Percentage string rounded to 1 decimal (e.g., "23.4%")
 */
function calculatePercentage(value, total) {
  if (total === 0) return '0.0%';
  return ((value / total) * 100).toFixed(1) + '%';
}

/**
 * Build datalabels config based on chart type.
 * @param {'doughnut'|'pie'|'bar'|'line'} chartType - Type of chart
 * @param {number[]} data - Dataset values
 * @returns {Object} chartjs-plugin-datalabels configuration object
 */
function buildLabelConfig(chartType, data) {
  const total = data.reduce((s, v) => s + Math.abs(v), 0);

  if (chartType === 'doughnut' || chartType === 'pie') {
    return {
      display: function(context) {
        return context.dataset.data[context.dataIndex] > 0;
      },
      formatter: function(value) {
        return calculatePercentage(Math.abs(value), total);
      },
      anchor: function(context) {
        const pct = (Math.abs(context.dataset.data[context.dataIndex]) / total) * 100;
        return pct >= 5 ? 'center' : 'end';
      },
      align: function(context) {
        const pct = (Math.abs(context.dataset.data[context.dataIndex]) / total) * 100;
        return pct >= 5 ? 'center' : 'end';
      },
      offset: function(context) {
        const pct = (Math.abs(context.dataset.data[context.dataIndex]) / total) * 100;
        return pct < 5 ? 10 : 0;
      },
      color: function(context) {
        const pct = (Math.abs(context.dataset.data[context.dataIndex]) / total) * 100;
        return pct >= 5 ? '#fff' : '#333';
      },
      font: { weight: 'bold', size: 11 }
    };
  }

  // Bar and Line charts
  return {
    display: function(context) {
      return context.dataset.data[context.dataIndex] !== 0;
    },
    formatter: function(value) {
      return typeof KPICalculator !== 'undefined' ? KPICalculator.formatBRL(value) : value.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' });
    },
    anchor: 'end',
    align: 'top',
    color: '#333',
    font: { size: 10, weight: 'bold' }
  };
}

// Expose for testing
if (typeof globalThis !== 'undefined') {
  globalThis.calculatePercentage = calculatePercentage;
  globalThis.buildLabelConfig = buildLabelConfig;
}

/**
 * Draw value labels directly on chart canvas for bar/line exports.
 * @param {object} chart - Chart.js instance
 * @param {string} type - Chart type
 * @param {number[]} values - Dataset values
 */
function drawLabelsOnCanvas(chart, type, values) {
  const ctx = chart.canvas.getContext('2d');
  const total = values.reduce((s, v) => s + Math.abs(v), 0);
  if (total === 0) return;

  const meta = chart.getDatasetMeta(0);
  if (!meta || !meta.data) return;

  ctx.save();

  if (type === 'doughnut' || type === 'pie') {
    meta.data.forEach((el, i) => {
      const value = Math.abs(values[i] || 0);
      if (value === 0) return;
      const pct = ((value / total) * 100).toFixed(1);
      const pctNum = (value / total) * 100;

      const midAngle = (el.startAngle + el.endAngle) / 2;
      const innerR = el.innerRadius || 0;
      const outerR = el.outerRadius || 0;

      if (pctNum >= 8) {
        const r = (innerR + outerR) / 2;
        const x = el.x + Math.cos(midAngle) * r;
        const y = el.y + Math.sin(midAngle) * r;

        ctx.font = 'bold 13px -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif';
        ctx.textAlign = 'center';
        ctx.textBaseline = 'middle';
        ctx.shadowColor = 'rgba(0, 0, 0, 0.6)';
        ctx.shadowBlur = 3;
        ctx.shadowOffsetX = 0;
        ctx.shadowOffsetY = 1;
        ctx.fillStyle = '#ffffff';
        ctx.fillText(pct + '%', x, y);
        ctx.shadowColor = 'transparent';
        ctx.shadowBlur = 0;
        ctx.shadowOffsetX = 0;
        ctx.shadowOffsetY = 0;
      } else if (pctNum >= 3) {
        const startR = outerR + 8;
        const endR = outerR + 28;
        const lineStartX = el.x + Math.cos(midAngle) * startR;
        const lineStartY = el.y + Math.sin(midAngle) * startR;
        const lineEndX = el.x + Math.cos(midAngle) * endR;
        const lineEndY = el.y + Math.sin(midAngle) * endR;

        ctx.beginPath();
        ctx.moveTo(lineStartX, lineStartY);
        ctx.lineTo(lineEndX, lineEndY);
        ctx.strokeStyle = '#6b7280';
        ctx.lineWidth = 1;
        ctx.stroke();

        const textX = el.x + Math.cos(midAngle) * (endR + 4);
        const textY = lineEndY;

        ctx.font = '600 11px -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif';
        ctx.textAlign = midAngle > Math.PI / 2 && midAngle < Math.PI * 1.5 ? 'right' : 'left';
        ctx.textBaseline = 'middle';

        const text = pct + '%';
        const textWidth = ctx.measureText(text).width;
        const pillX = ctx.textAlign === 'right' ? textX - textWidth - 6 : textX - 3;
        ctx.fillStyle = 'rgba(255, 255, 255, 0.9)';
        ctx.beginPath();
        ctx.roundRect(pillX, textY - 8, textWidth + 6, 16, 4);
        ctx.fill();
        ctx.strokeStyle = '#e2e8f0';
        ctx.lineWidth = 0.5;
        ctx.stroke();

        ctx.fillStyle = '#374151';
        ctx.fillText(text, textX, textY);
      }
    });
  } else {
    // Bar / Line charts
    ctx.font = '600 10px -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif';
    ctx.textAlign = 'center';
    ctx.textBaseline = 'bottom';

    meta.data.forEach((el, i) => {
      const value = values[i];
      if (!value || value === 0) return;
      const formatted = (typeof KPICalculator !== 'undefined')
        ? KPICalculator.formatBRL(value)
        : Math.abs(value).toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' });

      const x = el.x;
      const y = el.y - 8;

      const textWidth = ctx.measureText(formatted).width;
      ctx.fillStyle = 'rgba(255, 255, 255, 0.85)';
      ctx.beginPath();
      ctx.roundRect(x - textWidth / 2 - 4, y - 12, textWidth + 8, 14, 3);
      ctx.fill();

      ctx.fillStyle = '#1f2937';
      ctx.fillText(formatted, x, y);
    });
  }

  ctx.restore();
}

/**
 * Export a Chart.js instance as PNG.
 * For doughnut/pie charts: creates offscreen canvas with legend table below (no slice labels).
 * For bar/line charts: draws labels on canvas, exports, then refreshes chart.
 * @param {object} chartInstance - The Chart.js chart instance
 * @param {string} filename - Desired download filename (without extension)
 */
function exportChartWithLabels(chartInstance, filename) {
  if (!chartInstance || !chartInstance.canvas) {
    console.warn('exportChartWithLabels: chartInstance ou canvas não encontrado');
    return;
  }

  const chartType = chartInstance.config.type;
  const datasets = chartInstance.data.datasets;
  const data = datasets[0] ? datasets[0].data.map(v => Math.abs(Number(v) || 0)) : [];
  const isDoughnut = (chartType === 'doughnut' || chartType === 'pie');

  function triggerDownload(dataUrl, fname) {
    const link = document.createElement('a');
    link.style.display = 'none';
    link.download = (fname || 'chart') + '.png';
    link.href = dataUrl;
    document.body.appendChild(link);
    setTimeout(() => {
      link.click();
      setTimeout(() => { document.body.removeChild(link); }, 100);
    }, 0);
  }

  if (isDoughnut) {
    // Prepare legend entries
    const labels = chartInstance.data.labels || [];
    const colors = datasets[0].backgroundColor || [];
    const total = data.reduce((s, v) => s + v, 0);
    
    const entries = labels.map((name, i) => ({
      color: colors[i] || '#999',
      name: name || 'Sem nome',
      value: data[i],
      formattedValue: data[i].toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' }),
      percentage: total > 0 ? ((data[i] / total) * 100).toFixed(1) + '%' : '0.0%'
    })).sort((a, b) => b.value - a.value);

    // Calculate legend dimensions
    const ROW_HEIGHT = 24;
    const PADDING = 20;
    const LEGEND_HEIGHT = PADDING + (entries.length * ROW_HEIGHT) + PADDING;

    // Create offscreen canvas
    const chartCanvas = chartInstance.canvas;
    const chartWidth = chartCanvas.width;
    const chartHeight = chartCanvas.height;
    const offscreen = document.createElement('canvas');
    offscreen.width = chartWidth;
    offscreen.height = chartHeight + LEGEND_HEIGHT;
    const ctx = offscreen.getContext('2d');

    // White background
    ctx.fillStyle = '#ffffff';
    ctx.fillRect(0, 0, offscreen.width, offscreen.height);

    // Draw chart image (top portion)
    ctx.drawImage(chartCanvas, 0, 0, chartWidth, chartHeight);

    // Draw separator line
    ctx.strokeStyle = '#e2e8f0';
    ctx.lineWidth = 1;
    ctx.beginPath();
    ctx.moveTo(PADDING, chartHeight + 4);
    ctx.lineTo(chartWidth - PADDING, chartHeight + 4);
    ctx.stroke();

    // Draw legend table
    const SWATCH_SIZE = 12;
    const FONT = '12px -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif';
    const BOLD_FONT = 'bold 12px -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif';
    ctx.font = FONT;
    ctx.textBaseline = 'middle';

    let y = chartHeight + PADDING + 8;
    for (const entry of entries) {
      // Color swatch (rounded rect)
      ctx.fillStyle = entry.color;
      ctx.beginPath();
      ctx.roundRect(PADDING, y - SWATCH_SIZE / 2, SWATCH_SIZE, SWATCH_SIZE, 2);
      ctx.fill();

      // Category name
      ctx.fillStyle = '#1f2937';
      ctx.font = FONT;
      ctx.textAlign = 'left';
      const maxNameWidth = chartWidth * 0.45;
      let displayName = entry.name;
      if (ctx.measureText(displayName).width > maxNameWidth) {
        while (ctx.measureText(displayName + '...').width > maxNameWidth && displayName.length > 0) {
          displayName = displayName.slice(0, -1);
        }
        displayName += '...';
      }
      ctx.fillText(displayName, PADDING + SWATCH_SIZE + 10, y);

      // BRL value
      ctx.font = BOLD_FONT;
      ctx.textAlign = 'right';
      ctx.fillStyle = '#374151';
      ctx.fillText(entry.formattedValue, chartWidth - PADDING - 70, y);

      // Percentage
      ctx.font = FONT;
      ctx.fillStyle = '#6b7280';
      ctx.fillText(entry.percentage, chartWidth - PADDING, y);

      y += ROW_HEIGHT;
    }

    // Export offscreen canvas
    try {
      const dataUrl = offscreen.toDataURL('image/png', 1.0);
      triggerDownload(dataUrl, filename);
    } catch (err) {
      console.error('Erro ao exportar gráfico:', err);
    }
  } else {
    // Bar / Line charts: export with legend table below (same approach as doughnut)
    const labels = chartInstance.data.labels || [];
    const colors = datasets[0].backgroundColor || [];
    const total = data.reduce((s, v) => s + Math.abs(v), 0);

    const entries = labels.map((name, i) => ({
      color: Array.isArray(colors) ? (colors[i] || '#1d4ed8') : (colors || '#1d4ed8'),
      name: name || '',
      value: data[i],
      formattedValue: data[i].toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' }),
      percentage: total > 0 ? ((Math.abs(data[i]) / total) * 100).toFixed(1) + '%' : '0.0%'
    }));

    // Calculate legend dimensions
    const ROW_HEIGHT_BAR = 22;
    const PADDING_BAR = 20;
    const LEGEND_HEIGHT_BAR = PADDING_BAR + (entries.length * ROW_HEIGHT_BAR) + PADDING_BAR;

    const chartCanvas = chartInstance.canvas;
    const chartWidth = chartCanvas.width;
    const chartHeight = chartCanvas.height;
    const offscreen = document.createElement('canvas');
    offscreen.width = chartWidth;
    offscreen.height = chartHeight + LEGEND_HEIGHT_BAR;
    const ctx = offscreen.getContext('2d');

    // White background
    ctx.fillStyle = '#ffffff';
    ctx.fillRect(0, 0, offscreen.width, offscreen.height);

    // Draw chart image (top portion - clean, no labels on bars)
    ctx.drawImage(chartCanvas, 0, 0, chartWidth, chartHeight);

    // Draw separator line
    ctx.strokeStyle = '#e2e8f0';
    ctx.lineWidth = 1;
    ctx.beginPath();
    ctx.moveTo(PADDING_BAR, chartHeight + 4);
    ctx.lineTo(chartWidth - PADDING_BAR, chartHeight + 4);
    ctx.stroke();

    // Draw legend table
    const SWATCH_SIZE_BAR = 10;
    const FONT_BAR = '11px -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif';
    const BOLD_FONT_BAR = 'bold 11px -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif';
    ctx.textBaseline = 'middle';

    let yBar = chartHeight + PADDING_BAR + 6;
    for (const entry of entries) {
      // Color swatch
      ctx.fillStyle = entry.color;
      ctx.beginPath();
      ctx.roundRect(PADDING_BAR, yBar - SWATCH_SIZE_BAR / 2, SWATCH_SIZE_BAR, SWATCH_SIZE_BAR, 2);
      ctx.fill();

      // Label name
      ctx.fillStyle = '#1f2937';
      ctx.font = FONT_BAR;
      ctx.textAlign = 'left';
      ctx.fillText(entry.name, PADDING_BAR + SWATCH_SIZE_BAR + 8, yBar);

      // BRL value
      ctx.font = BOLD_FONT_BAR;
      ctx.textAlign = 'right';
      ctx.fillStyle = '#374151';
      ctx.fillText(entry.formattedValue, chartWidth - PADDING_BAR, yBar);

      yBar += ROW_HEIGHT_BAR;
    }

    // Export offscreen canvas
    try {
      const dataUrl = offscreen.toDataURL('image/png', 1.0);
      triggerDownload(dataUrl, filename);
    } catch (err) {
      console.error('Erro ao exportar gráfico:', err);
    }
  }
}

// Expose for testing
if (typeof globalThis !== 'undefined') {
  globalThis.exportChartWithLabels = exportChartWithLabels;
}

// ============================================================
// SECTION: Chart Image Export & Fullscreen Mode
// ============================================================

/**
 * Wire per-chart export buttons to exportChartWithLabels (with datalabels during export).
 * Wire fullscreen mode toggle via #btn-fullscreen.
 * Validates: Requirements 2.1, 2.2, 2.6, 2.7, 8.2, 8.3
 */
(function initChartExportAndFullscreen() {
  function setup() {
    // --- Per-chart export buttons ---
    const chartExportButtons = document.querySelectorAll('.btn-chart-export');
    chartExportButtons.forEach(btn => {
      btn.addEventListener('click', (e) => {
        e.preventDefault();
        e.stopPropagation();
        const chartId = btn.getAttribute('data-chart-id');
        if (!chartId) return;

        const chartInstance = ChartEngine._instances[chartId];
        if (!chartInstance) {
          console.warn(`Chart "${chartId}" não encontrado. Instâncias disponíveis:`, Object.keys(ChartEngine._instances));
          return;
        }

        const filename = `grafico-${chartId}-${new Date().toISOString().slice(0, 10)}`;

        // Check if receitas/despesas chart has expanded panel visible
        if (chartId === 'receitas-chart' || chartId === 'despesas-chart') {
          const expandType = chartId === 'receitas-chart' ? 'receitas' : 'despesas';
          const panel = document.getElementById(`expand-${expandType}`);

          if (panel && !panel.hidden && !panel.hasAttribute('hidden')) {
            const nivel2Canvas = document.getElementById(`chart-${expandType}-nivel2`);
            const nivel3Canvas = document.getElementById(`chart-${expandType}-nivel3`);

            const mainCanvas = chartInstance.canvas;
            const canvases = [mainCanvas];
            if (nivel2Canvas) canvases.push(nivel2Canvas);
            if (nivel3Canvas && nivel3Canvas.style.display !== 'none') canvases.push(nivel3Canvas);

            const maxWidth = Math.max(...canvases.map(c => c.width));
            const totalHeight = canvases.reduce((h, c) => h + c.height + 20, 0) + 20;

            const offscreen = document.createElement('canvas');
            offscreen.width = maxWidth;
            offscreen.height = totalHeight;
            const ctx = offscreen.getContext('2d');
            ctx.fillStyle = '#ffffff';
            ctx.fillRect(0, 0, offscreen.width, offscreen.height);

            let yOffset = 10;
            for (const canvas of canvases) {
              const x = (maxWidth - canvas.width) / 2;
              ctx.drawImage(canvas, x, yOffset, canvas.width, canvas.height);
              yOffset += canvas.height + 20;
            }

            try {
              const dataUrl = offscreen.toDataURL('image/png', 1.0);
              const link = document.createElement('a');
              link.style.display = 'none';
              link.download = filename + '.png';
              link.href = dataUrl;
              document.body.appendChild(link);
              setTimeout(() => { link.click(); setTimeout(() => { document.body.removeChild(link); }, 100); }, 0);
            } catch (err) {
              console.error('Erro ao exportar gráficos expandidos:', err);
            }
            return;
          }
        }

        exportChartWithLabels(chartInstance, filename);
      });
    });

    // --- Chart expand buttons (Receitas / Despesas detail) ---
    const _expandCharts = {};
    document.querySelectorAll('.btn-chart-expand').forEach(btn => {
      btn.addEventListener('click', () => {
        const expandType = btn.dataset.expand; // 'receitas' or 'despesas'
        const panel = document.getElementById(`expand-${expandType}`);
        if (!panel) return;

        const isOpen = !panel.hidden;
        panel.hidden = isOpen;
        btn.classList.toggle('active', !isOpen);
        btn.innerHTML = isOpen ? '&#128269; Detalhar' : '&#10006; Fechar';

        if (!isOpen) {
          // Render detail charts
          const data = UIController._currentData || DataLayer.getData();
          const filtered = expandType === 'receitas'
            ? data.filter(r => r.tipo === 'receita')
            : data.filter(r => r.tipo === 'despesa');

          // Nivel 2 chart (nivel1 field = 2ª Categoria)
          const nivel2Canvas = document.getElementById(`chart-${expandType}-nivel2`);
          if (nivel2Canvas) {
            if (_expandCharts[`${expandType}_n2`]) _expandCharts[`${expandType}_n2`].destroy();
            const totals2 = {};
            for (const r of filtered) {
              const cat = r.nivel1 || 'Sem Categoria';
              totals2[cat] = (totals2[cat] || 0) + r.valor;
            }
            const sorted2 = Object.entries(totals2).sort((a, b) => b[1] - a[1]).slice(0, 10);
            const colors = ChartEngine._colors || ['#4f46e5','#06b6d4','#8b5cf6','#f59e0b','#10b981','#ef4444','#ec4899','#6366f1','#14b8a6','#f97316'];
            _expandCharts[`${expandType}_n2`] = new Chart(nivel2Canvas, {
              type: 'doughnut',
              data: {
                labels: sorted2.map(([n]) => n),
                datasets: [{ data: sorted2.map(([,v]) => v), backgroundColor: colors.slice(0, sorted2.length), borderColor: '#fff', borderWidth: 2 }]
              },
              options: { responsive: true, plugins: { legend: { position: 'bottom', labels: { boxWidth: 10, font: { size: 10 } } }, tooltip: { callbacks: { label: (ctx) => `${ctx.label}: ${KPICalculator.formatBRL(ctx.raw)}` } } } }
            });
          }

          // Nivel 3 chart (nivel2 field = 3ª Categoria) - Pie Chart
          const nivel3Canvas = document.getElementById(`chart-${expandType}-nivel3`);
          if (nivel3Canvas) {
            if (_expandCharts[`${expandType}_n3`]) _expandCharts[`${expandType}_n3`].destroy();

            // Get pie data using _preparePieChartData aggregation
            const pieData = _preparePieChartData(filtered);

            if (!pieData) {
              // Empty state handling
              nivel3Canvas.style.display = 'none';
              // Remove any existing placeholder
              const existingPlaceholder = nivel3Canvas.parentNode.querySelector('.chart-empty-placeholder');
              if (existingPlaceholder) existingPlaceholder.remove();
              // Add placeholder message
              const placeholder = document.createElement('p');
              placeholder.className = 'chart-empty-placeholder';
              placeholder.textContent = 'Nenhum dado de 3ª categoria disponível para os filtros selecionados.';
              nivel3Canvas.parentNode.appendChild(placeholder);
            } else {
              nivel3Canvas.style.display = '';
              // Remove placeholder if exists
              const existingPlaceholder = nivel3Canvas.parentNode.querySelector('.chart-empty-placeholder');
              if (existingPlaceholder) existingPlaceholder.remove();

              _expandCharts[`${expandType}_n3`] = new Chart(nivel3Canvas, {
                type: 'pie',
                data: {
                  labels: pieData.labels,
                  datasets: [{
                    data: pieData.data,
                    backgroundColor: pieData.backgroundColor,
                    borderColor: '#fff',
                    borderWidth: 2
                  }]
                },
                options: {
                  responsive: true,
                  maintainAspectRatio: true,
                  plugins: {
                    legend: {
                      position: 'bottom',
                      labels: {
                        boxWidth: 12,
                        font: { size: 11 },
                        padding: 8
                      }
                    },
                    tooltip: {
                      callbacks: {
                        label: (ctx) => {
                          const total = ctx.dataset.data.reduce((s, v) => s + v, 0);
                          const pct = ((ctx.raw / total) * 100).toFixed(1);
                          const formatted = ctx.raw.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' });
                          return `${ctx.label}: ${formatted} (${pct}%)`;
                        }
                      }
                    }
                  }
                }
              });
            }
          }
        } else {
          // Destroy expanded charts on close
          if (_expandCharts[`${expandType}_n2`]) { _expandCharts[`${expandType}_n2`].destroy(); _expandCharts[`${expandType}_n2`] = null; }
          if (_expandCharts[`${expandType}_n3`]) { _expandCharts[`${expandType}_n3`].destroy(); _expandCharts[`${expandType}_n3`] = null; }
        }
      });
    });

    // --- Chart reset buttons (clear all dashboard filters) ---
    document.querySelectorAll('.btn-chart-reset').forEach(btn => {
      btn.addEventListener('click', () => {
        UIController.clearFilters();
      });
    });

    // --- KPI card click (opens KPIDetailModal) ---
    document.addEventListener('click', (e) => {
      const card = e.target.closest('.kpi-card[data-kpi]');
      if (!card) return;
      const kpiKey = card.dataset.kpi;
      if (kpiKey && typeof KPIDetailModal !== 'undefined') {
        KPIDetailModal.open(kpiKey);
      }
    });

    // --- Fullscreen mode toggle ---
    const btnFullscreen = document.getElementById('btn-fullscreen');
    if (btnFullscreen) {
      btnFullscreen.addEventListener('click', () => {
        const isFullscreen = document.body.classList.toggle('fullscreen');

        // Update button text
        btnFullscreen.innerHTML = isFullscreen
          ? '<span aria-hidden="true">&#9974;</span> Sair Tela Cheia'
          : '<span aria-hidden="true">&#9974;</span> Tela Cheia';

        // Update aria-label
        btnFullscreen.setAttribute('aria-label', isFullscreen
          ? 'Sair do modo tela cheia'
          : 'Ativar modo tela cheia');

        // Use Fullscreen API if available
        if (isFullscreen) {
          if (document.documentElement.requestFullscreen) {
            document.documentElement.requestFullscreen().catch(() => {
              // Fullscreen API not supported or denied - CSS class still handles layout
            });
          }
        } else {
          if (document.fullscreenElement && document.exitFullscreen) {
            document.exitFullscreen().catch(() => {});
          }
        }
      });

      // Listen for fullscreen exit via Escape key (browser native)
      document.addEventListener('fullscreenchange', () => {
        if (!document.fullscreenElement && document.body.classList.contains('fullscreen')) {
          document.body.classList.remove('fullscreen');
          btnFullscreen.innerHTML = '<span aria-hidden="true">&#9974;</span> Tela Cheia';
          btnFullscreen.setAttribute('aria-label', 'Ativar modo tela cheia');
        }
      });
    }
  }

  // --- Categorias Modal: open/close and render charts ---
  let _nivel2ChartInstance = null;
  let _nivel3ChartInstance = null;

  function openCategoriasModal() {
    const modal = document.getElementById('categorias-modal');
    if (modal) modal.hidden = false;
    renderCategoriasCharts();
  }

  function closeCategoriasModal() {
    const modal = document.getElementById('categorias-modal');
    if (modal) modal.hidden = true;
  }

  function renderCategoriasCharts() {
    const data = DataLayer.getFilteredData(DataLayer.getActiveFilters());

    // Render Nivel 2 chart
    const nivel2Data = ChartEngine._prepareNivel2Chart(data);
    const canvas2 = document.getElementById('chart-nivel2-chart');
    if (canvas2) {
      if (_nivel2ChartInstance) _nivel2ChartInstance.destroy();
      _nivel2ChartInstance = new Chart(canvas2.getContext('2d'), {
        type: 'doughnut',
        data: nivel2Data,
        options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { position: 'right' } }, cutout: '50%', animation: { duration: 750 } }
      });
    }

    // Render Nivel 3 chart
    const nivel3Data = ChartEngine._prepareNivel3Chart(data);
    const canvas3 = document.getElementById('chart-nivel3-chart');
    if (canvas3) {
      if (_nivel3ChartInstance) _nivel3ChartInstance.destroy();
      _nivel3ChartInstance = new Chart(canvas3.getContext('2d'), {
        type: 'bar',
        data: nivel3Data,
        options: { responsive: true, maintainAspectRatio: false, indexAxis: 'y', animation: { duration: 750 }, plugins: { legend: { display: false } }, scales: { x: { beginAtZero: true } } }
      });
    }
  }

  // --- Detail Tabs ---
  function setupDetailTabs() {
    const tabs = document.querySelectorAll('.detail-tab');
    const contents = document.querySelectorAll('.detail-tab-content');

    tabs.forEach(tab => {
      tab.addEventListener('click', () => {
        tabs.forEach(t => t.classList.remove('active'));
        contents.forEach(c => c.classList.remove('active'));
        tab.classList.add('active');
        const target = document.getElementById('tab-' + tab.dataset.tab);
        if (target) target.classList.add('active');
      });
    });

    // Conta filter for Dados Detalhados
    const contaFilter = document.getElementById('detail-conta-filter');
    if (contaFilter) {
      // Populate on data load
      EventBus.on('data:loaded', (payload) => {
        const contas = [...new Set(payload.records.map(r => r.conta).filter(c => c))].sort();
        contaFilter.innerHTML = '<option value="all">Todas</option>';
        contas.forEach(c => { contaFilter.innerHTML += `<option value="${c}">${c}</option>`; });
      });

      // Filter tables when conta changes
      contaFilter.addEventListener('change', () => {
        const value = contaFilter.value;
        const allData = DataLayer.getData();
        const filtered = value === 'all' ? allData : allData.filter(r => r.conta === value);
        UIController.initTables(filtered);
      });
    }
  }

  function setupCategoriasModal() {
    const btnOpen = document.getElementById('btn-open-categorias');
    const btnClose = document.getElementById('btn-close-categorias');
    const nivel2Filter = document.getElementById('nivel2-filter');
    const nivel3Filter = document.getElementById('nivel3-filter');

    if (btnOpen) btnOpen.addEventListener('click', openCategoriasModal);
    if (btnClose) btnClose.addEventListener('click', closeCategoriasModal);

    // Close on overlay click
    const modal = document.getElementById('categorias-modal');
    if (modal) {
      modal.addEventListener('click', (e) => {
        if (e.target === modal) closeCategoriasModal();
      });
    }

    // Dropdown filters inside modal
    if (nivel2Filter) {
      nivel2Filter.addEventListener('change', () => {
        const data = DataLayer.getFilteredData(DataLayer.getActiveFilters());
        const nivel2Data = ChartEngine._prepareNivel2Chart(data);
        if (_nivel2ChartInstance) {
          _nivel2ChartInstance.data = nivel2Data;
          _nivel2ChartInstance.update('active');
        }
      });
    }
    if (nivel3Filter) {
      nivel3Filter.addEventListener('change', () => {
        const data = DataLayer.getFilteredData(DataLayer.getActiveFilters());
        const nivel3Data = ChartEngine._prepareNivel3Chart(data);
        if (_nivel3ChartInstance) {
          _nivel3ChartInstance.data = nivel3Data;
          _nivel3ChartInstance.update('active');
        }
      });
    }
  }

  // --- Ranking Selector ---
  function setupRankingSelector() {
    const selector = document.getElementById('ranking-selector');
    if (!selector) return;

    selector.addEventListener('change', () => {
      const groups = ['projetos', 'clientes', 'categorias'];
      groups.forEach(g => {
        const el = document.getElementById('ranking-group-' + g);
        if (el) el.hidden = (g !== selector.value);
      });
    });
  }

  // Initialize when DOM is ready
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', () => { setup(); setupCategoriasModal(); setupDetailTabs(); setupRankingSelector(); });
  } else {
    setup();
    setupCategoriasModal();
    setupDetailTabs();
    setupRankingSelector();
    setupCategoriasModal();
  }
})();

// ============================================================
// SECTION: KPI Card Click - Standalone Handler (fallback)
// ============================================================
(function() {
  function initKPIClicks() {
    document.addEventListener('click', function(e) {
      var card = e.target.closest ? e.target.closest('.kpi-card[data-kpi]') : null;
      if (!card) return;
      var kpiKey = card.getAttribute('data-kpi');
      if (!kpiKey) return;
      if (typeof KPIDetailModal !== 'undefined' && KPIDetailModal.open) {
        KPIDetailModal.open(kpiKey);
      }
    });
  }
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', initKPIClicks);
  } else {
    initKPIClicks();
  }
})();

// ============================================================
// SECTION: Module Exports (for testing in Node.js environment)
// ============================================================
/* istanbul ignore next */
if (typeof globalThis !== 'undefined') {
  globalThis.computePageButtons = computePageButtons;
}
