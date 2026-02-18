import React, { useState, useEffect, useRef, useCallback } from 'react'
import { useParams, useNavigate } from 'react-router-dom'
import { SpreadsheetComponent, SheetsDirective, SheetDirective, ColumnsDirective, ColumnDirective } from '@syncfusion/ej2-react-spreadsheet'
import { FiEye, FiSettings, FiArrowLeft, FiCheckSquare } from 'react-icons/fi'
import toast from 'react-hot-toast'
import Sidebar from '../components/Sidebar'
import TopBar from '../components/TopBar'
import RawDataPreviewModal from '../components/RawDataPreviewModal'
import UnusedRawDataModal from '../components/UnusedRawDataModal'
import ProposalSettingsModal from '../components/ProposalSettingsModal'
import { proposalAPI } from '../services/proposalService'
import { generateCalculationSheet, generateColumnConfigs } from '../utils/generateCalculationSheet'
import { generateDemolitionFormulas } from '../utils/processors/demolitionProcessor'
import { generateExcavationFormulas } from '../utils/processors/excavationProcessor'
import { generateRockExcavationFormulas } from '../utils/processors/rockExcavationProcessor'
import { generateSoeFormulas } from '../utils/processors/soeProcessor'
import { generateFoundationFormulas } from '../utils/processors/foundationProcessor'
import { generateWaterproofingFormulas } from '../utils/processors/waterproofingProcessor'
import { buildProposalSheet } from '../utils/buildProposalSheet'
import { useSidebar } from '../context/SidebarContext'

const ProposalDetail = () => {
  const { proposalId } = useParams()
  const navigate = useNavigate()
  const { sidebarCollapsed, toggleSidebar } = useSidebar()
  const [proposal, setProposal] = useState(null)
  const [isLoading, setIsLoading] = useState(true)
  const [calculationData, setCalculationData] = useState([])
  const [formulaData, setFormulaData] = useState([])
  const [rockExcavationTotals, setRockExcavationTotals] = useState({ totalSQFT: 0, totalCY: 0 })
  const [lineDrillTotalFT, setLineDrillTotalFT] = useState(0)
  const rawDataRef = useRef(null)
  const proposalBuiltRef = useRef(false)
  const calculationDataRef = useRef([])
  const generatedDataRef = useRef(null)
  const formulaDataRef = useRef([])
  const rockExcavationTotalsRef = useRef({ totalSQFT: 0, totalCY: 0 })
  const lineDrillTotalFTRef = useRef(0)
  const unusedRawDataRowsRef = useRef([])
  calculationDataRef.current = calculationData
  formulaDataRef.current = formulaData
  rockExcavationTotalsRef.current = rockExcavationTotals
  lineDrillTotalFTRef.current = lineDrillTotalFT
  const [isPreviewModalOpen, setIsPreviewModalOpen] = useState(false)
  const [isUnusedDataModalOpen, setIsUnusedDataModalOpen] = useState(false)
  const [isSettingsModalOpen, setIsSettingsModalOpen] = useState(false)
  const [isSaving, setIsSaving] = useState(false)
  const [lastSaved, setLastSaved] = useState(null)
  const [isSpreadsheetLoading, setIsSpreadsheetLoading] = useState(false)
  const [hasUnsavedChanges, setHasUnsavedChanges] = useState(false)
  const [saveError, setSaveError] = useState(null)
  const spreadsheetRef = useRef(null)
  const saveTimeoutRef = useRef(null)
  const maxWaitTimeoutRef = useRef(null)
  const hasLoadedFromJson = useRef(false)
  const needReapplyAfterRawSave = useRef(false)
  const isSavingRef = useRef(false)
  const pendingSaveRef = useRef(false)
  const retryCountRef = useRef(0)
  const lastSaveTimeRef = useRef(null)

  // Load proposal data
  useEffect(() => {
    if (proposalId) {
      fetchProposal()
    }
  }, [proposalId])

  const fetchProposal = async () => {
    try {
      setIsLoading(true)
      const response = await proposalAPI.getById(proposalId)
      setProposal(response.proposal)
    } catch (error) {
      console.error('Error fetching proposal:', error)
      toast.error('Error loading proposal')
      navigate('/proposals')
    } finally {
      setIsLoading(false)
    }
  }

  // Generate calculation data from rawExcelData when available (needed for buildProposalSheet)
  // Run for both new proposals (no spreadsheetJson) and saved proposals (have spreadsheetJson + rawExcelData from API)
  useEffect(() => {
    if (!proposal || !proposal.rawExcelData) {
      generatedDataRef.current = null
      return
    }

    proposalBuiltRef.current = false
    const { headers, rows } = proposal.rawExcelData
    const rawData = [headers, ...rows]
    rawDataRef.current = rawData

    const template = proposal.template || 'capstone'
    const result = generateCalculationSheet(template, rawData)
    setCalculationData(result.rows)
    setFormulaData(result.formulas)
    setRockExcavationTotals(result.rockExcavationTotals || { totalSQFT: 0, totalCY: 0 })
    setLineDrillTotalFT(result.lineDrillTotalFT || 0)

    setLineDrillTotalFT(result.lineDrillTotalFT || 0)

    // Handle unused raw data rows
    const calculatedUnusedRows = result.unusedRawDataRows || []
    const currentDbRows = proposal.unusedRawDataRows || []
    const dbRowMap = new Map(currentDbRows.map(r => [r.rowIndex, r]))

    // Merge: use calculated rows but preserve isUsed status from DB
    const mergedUnusedRows = calculatedUnusedRows.map(calcRow => {
      const dbRow = dbRowMap.get(calcRow.rowIndex)
      return {
        ...calcRow,
        isUsed: dbRow ? dbRow.isUsed : false
      }
    })

    unusedRawDataRowsRef.current = mergedUnusedRows

    // Synch to proposal state if different (initial load or new calculation)
    // Check if lengths differ or if any rowIndex is missing/new
    const isDifferent = mergedUnusedRows.length !== currentDbRows.length ||
      mergedUnusedRows.some((r, i) => r.rowIndex !== currentDbRows[i]?.rowIndex)

    if (isDifferent) {
      setProposal(prev => ({ ...prev, unusedRawDataRows: mergedUnusedRows }))
    }

    // Store window globals for proposal sheet (used by buildProposalSheet)
    window.soldierPileGroups = result.soldierPileGroups || []
    window.soeSubsectionItems = new Map()
    window.primarySecantItems = result.primarySecantItems || []
    window.secondarySecantItems = result.secondarySecantItems || []
    window.tangentPileItems = result.tangentPileItems || []
    window.pargingItems = result.pargingItems || []
    window.guideWallItems = result.guideWallItems || []
    window.dowelBarItems = result.dowelBarItems || []
    window.rockPinItems = result.rockPinItems || []
    window.rockStabilizationItems = result.rockStabilizationItems || []
    window.buttonItems = result.buttonItems || []
    window.drilledFoundationPileGroups = result.drilledFoundationPileGroups || []
    window.helicalFoundationPileGroups = result.helicalFoundationPileGroups || []
    window.drivenFoundationPileItems = result.drivenFoundationPileItems || []
    window.stelcorDrilledDisplacementPileItems = result.stelcorDrilledDisplacementPileItems || []
    window.cfaPileItems = result.cfaPileItems || []
    window.foundationSubsectionItems = new Map()
    window.shotcreteItems = result.shotcreteItems || []
    window.permissionGroutingItems = result.permissionGroutingItems || []
    window.mudSlabItems = result.mudSlabItems || []

    // Check which waterproofing items are present
    if (rawData && rawData.length > 1) {
      const digitizerIdx = headers.findIndex(h => h && String(h).toLowerCase().trim() === 'digitizer item')
      const estimateIdx = headers.findIndex(h => h && String(h).toLowerCase().trim() === 'estimate')
      const itemsToCheck = {
        'foundation walls': /^FW\s*\(/i, 'retaining wall': /^RW\s*\(/i, 'vehicle barrier wall': /vehicle\s+barrier\s+wall\s*\(/i,
        'concrete liner wall': /concrete\s+liner\s+wall\s*\(/i, 'stem wall': /stem\s+wall\s*\(/i, 'grease trap pit wall': /grease\s+trap\s+pit\s+wall/i,
        'house trap pit wall': /house\s+trap\s+pit\s+wall/i, 'detention tank wall': /detention\s+tank\s+wall/i,
        'elevator pit walls': /elev(?:ator)?\s+pit\s+wall/i, 'duplex sewage ejector pit wall': /duplex\s+sewage\s+ejector\s+pit\s+wall/i
      }
      const negativeSideItemsToCheck = {
        'detention tank wall': /detention\s+tank\s+wall/i, 'elevator pit walls': /elev(?:ator)?\s+pit\s+wall/i,
        'detention tank slab': /detention\s+tank\s+slab(?!\s+lid)/i, 'duplex sewage ejector pit wall': /duplex\s+sewage\s+ejector\s+pit\s+wall/i,
        'duplex sewage ejector pit slab': /duplex\s+sewage\s+ejector\s+pit\s+slab/i, 'elevator pit slab': /elev(?:ator)?\s+pit\s+slab/i
      }
      const presentItems = []
      const presentNegativeSideItems = []
      const dataRows = rawData.slice(1)
      dataRows.forEach(row => {
        const digitizerItem = row[digitizerIdx]
        if (!digitizerItem) return
        if (estimateIdx >= 0 && row[estimateIdx] && String(row[estimateIdx]).trim() !== 'Waterproofing') return
        const itemText = String(digitizerItem).trim()
        Object.entries(itemsToCheck).forEach(([itemName, pattern]) => {
          if (pattern.test(itemText) && !presentItems.includes(itemName)) presentItems.push(itemName)
        })
        Object.entries(negativeSideItemsToCheck).forEach(([itemName, pattern]) => {
          if (pattern.test(itemText) && !presentNegativeSideItems.includes(itemName)) presentNegativeSideItems.push(itemName)
        })
      })
      window.waterproofingPresentItems = presentItems
      window.waterproofingNegativeSideItems = presentNegativeSideItems
    }

    // Store for buildProposalSheet when loading from JSON (state may not have updated yet)
    generatedDataRef.current = {
      rows: result.rows,
      formulas: result.formulas,
      rockExcavationTotals: result.rockExcavationTotals || { totalSQFT: 0, totalCY: 0 },
      lineDrillTotalFT: result.lineDrillTotalFT || 0,
      unusedRawDataRows: mergedUnusedRows
    }
  }, [proposal?.rawExcelData, proposal?.template])

  // Apply data and formulas or load from JSON
  useEffect(() => {
    if (!spreadsheetRef.current || !proposal) return

    // When we have rawExcelData, ONLY ever build from raw (never load spreadsheet from DB)
    if (proposal.rawExcelData) {
      if (needReapplyAfterRawSave.current && calculationData.length > 0 && hasLoadedFromJson.current) {
        needReapplyAfterRawSave.current = false
        applyDataToSpreadsheet().then(() => {
          setTimeout(() => {
            try { spreadsheetRef.current?.goTo('Calculations Sheet!A1') } catch (e) { }
          }, 100)
        })
        return
      }
      if (calculationData.length > 0 && !hasLoadedFromJson.current) {
        applyDataToSpreadsheet()
      }
      return
    }

    // Load from saved JSON only when there is no rawExcelData (e.g. legacy proposals)
    if (proposal.spreadsheetJson && !hasLoadedFromJson.current) {
      setIsSpreadsheetLoading(true)
        ; (async () => {
          try {
            const jsonData = proposal.spreadsheetJson.Workbook
              ? proposal.spreadsheetJson
              : { Workbook: proposal.spreadsheetJson }
            spreadsheetRef.current.openFromJson({ file: jsonData })
            hasLoadedFromJson.current = true
            setLastSaved(new Date(proposal.updatedAt))
            if (proposal.images && proposal.images.length > 0) {
              restoreImages(proposal.images)
            }
          } catch (error) {
            toast.error('Error loading saved spreadsheet')
          } finally {
            setIsSpreadsheetLoading(false)
          }
        })()
    }
    // Re-apply after save/rebuild when there is no rawExcelData (legacy path)
    else if (needReapplyAfterRawSave.current && calculationData.length > 0 && hasLoadedFromJson.current && spreadsheetRef.current) {
      needReapplyAfterRawSave.current = false
      applyDataToSpreadsheet().then(() => {
        setTimeout(() => {
          try { spreadsheetRef.current?.goTo('Calculations Sheet!A1') } catch (e) { }
        }, 100)
      })
    }
  }, [calculationData, formulaData, proposal])

  const applyDataToSpreadsheet = async () => {
    if (!spreadsheetRef.current) return

    setIsSpreadsheetLoading(true)
    // Use setTimeout to allow React to update the UI with loading state
    await new Promise(resolve => setTimeout(resolve, 50))

    try {
      // Build complete spreadsheet model as JSON for batch loading
      const spreadsheetModel = buildSpreadsheetModel()

      // Load the entire spreadsheet at once - much faster than individual cell updates
      spreadsheetRef.current.openFromJson({ file: spreadsheetModel })
      hasLoadedFromJson.current = true

      // Apply all formulas and styles after data is loaded
      // Small delay to ensure spreadsheet is fully rendered
      await new Promise(resolve => setTimeout(resolve, 100))
      applyFormulasAndStyles()

      // Build the Proposal sheet from calculation data
      try {
        buildProposalSheet(spreadsheetRef.current, {
          calculationData,
          formulaData,
          rockExcavationTotals,
          lineDrillTotalFT,
          rawData: rawDataRef.current
        })
        proposalBuiltRef.current = true
      } catch (e) {
        console.error('Error building proposal sheet:', e)
      }

      // Trigger initial save to persist both Calculations and Proposal sheets
      markDirtyAndScheduleSave()
      // Immediate save so Proposal Sheet content persists (programmatic updates may not trigger save events)
      await new Promise(resolve => setTimeout(resolve, 300))
      if (saveTimeoutRef.current) {
        clearTimeout(saveTimeoutRef.current)
        saveTimeoutRef.current = null
      }
      saveSpreadsheet(false)
    } catch (error) {
      console.error('Error loading spreadsheet model:', error)
      toast.error('Error loading spreadsheet')
    } finally {
      setIsSpreadsheetLoading(false)
    }
  }

  // Build spreadsheet model with raw data only (formulas and styles applied after loading)
  const buildSpreadsheetModel = () => {
    const calculationsRows = []

    // Build rows for Calculations sheet - raw data only
    calculationData.forEach((row, rowIndex) => {
      const cells = []

      row.forEach((cellValue, colIndex) => {
        const cell = {}

        if (cellValue !== '' && cellValue !== null && cellValue !== undefined) {
          // For Particulars column (B), prevent date conversion
          if (colIndex === 1 && typeof cellValue === 'string') {
            if (/^[A-Z]+-\d+/.test(cellValue) || /^\d+-\d+/.test(cellValue)) {
              cell.value = "'" + cellValue
            } else {
              cell.value = cellValue
            }
            cell.format = '@' // Text format
          } else {
            cell.value = cellValue
          }
        }

        cells.push(cell)
      })

      calculationsRows.push({ cells })
    })

    // Column configurations
    const columns = generateColumnConfigs().map(config => ({ width: config.width }))

    // Build the complete workbook model (no formulas - they're applied after loading)
    // Sheet order: Proposal Sheet first, then Calculations Sheet (names must match buildProposalSheet)
    const workbookModel = {
      Workbook: {
        sheets: [
          {
            name: 'Proposal Sheet',
            rows: [],
            columns: []
          },
          {
            name: 'Calculations Sheet',
            rows: calculationsRows,
            columns: columns
          }
        ],
        activeSheetIndex: 1
      }
    }

    return workbookModel
  }

  // NOTE: The complete formula and styling logic from Spreadsheet.jsx needs to be implemented here.
  // Due to the size (~1900 lines), this is a reference to copy the logic from:
  // frontend/src/pages/Spreadsheet.jsx lines 81-2169
  // 
  // The applyFormulasAndStyles function should:
  // 1. Apply formulas using spreadsheet.updateCell({ formula: ... })
  // 2. Apply styles using spreadsheet.cellFormat({...})
  // 3. Handle all sections: demolition, excavation, rock_excavation, soe, foundation, waterproofing, superstructure, trenching
  // 
  // For now, using a simplified version - copy the full logic from Spreadsheet.jsx for production
  const applyFormulasAndStyles = () => {
    if (!spreadsheetRef.current) return
    const spreadsheet = spreadsheetRef.current

    // Ensure we're on the Calculations Sheet when applying formulas
    try {
      spreadsheet.goTo('Calculations Sheet!A1')
    } catch (e) {
      // ignore
    }

    // Helper function to determine column B color based on takeoff value
    const getColumnBColor = (row, parsedData) => {
      // Check calculationData for the initial value in column C
      const rowIndex = row - 1 // calculationData is 0-indexed
      if (rowIndex >= 0 && rowIndex < calculationData.length) {
        const takeoffValue = calculationData[rowIndex][2] // Column C (index 2)
        
        // Consider it as "has data" (RED) only if:
        // 1. Not null/undefined
        // 2. Not empty string or whitespace
        // 3. Not zero (0 or '0')
        // 4. Not a formula (starts with =)
        // 5. Parses to a number > 0
        
        if (takeoffValue == null || takeoffValue === '' || takeoffValue === 0 || takeoffValue === '0') {
          // Definitely no data - use black
          return '#000000'
        }
        
        const valueStr = String(takeoffValue).trim()
        
        if (valueStr === '' || valueStr.startsWith('=')) {
          // Empty or formula reference - use black
          return '#000000'
        }
        
        const numValue = parseFloat(valueStr)
        if (!isNaN(numValue)) {
          // It's a number - only red if > 0
          return numValue > 0 ? '#FF0000' : '#000000'
        }
        
        // It's a non-numeric non-empty string - consider as having data
        return '#FF0000'
      }
      
      // Fallback: check parsedData.takeoff (for items coming from raw data)
      if (parsedData && parsedData.takeoff != null) {
        const takeoff = parsedData.takeoff
        
        if (takeoff === '' || takeoff === 0 || takeoff === '0') {
          return '#000000'
        }
        
        const takeoffStr = String(takeoff).trim()
        if (takeoffStr === '' || takeoffStr.startsWith('=')) {
          return '#000000'
        }
        
        const numValue = parseFloat(takeoffStr)
        if (!isNaN(numValue)) {
          return numValue > 0 ? '#FF0000' : '#000000'
        }
        
        // Non-numeric non-empty string
        return '#FF0000'
      }
      
      return '#000000' // Black if no data found
    }

    // Deferred foundation sum formulas
    const deferredFoundationSumL = []

    // Apply formulas from formulaData
    formulaData.forEach((formulaInfo) => {
      const { row, itemType, parsedData, section, subsection } = formulaInfo

      let formulas
      try {
        if (section === 'demolition') {
          if (itemType === 'demolition_sum') {
            const { firstDataRow, lastDataRow, subsection: subName } = formulaInfo
            spreadsheet.updateCell({ formula: `=SUM(J${firstDataRow}:J${lastDataRow})` }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `J${row}`)
            spreadsheet.updateCell({ formula: `=SUM(L${firstDataRow}:L${lastDataRow})` }, `L${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `L${row}`)
            if (subName === 'Demo isolated footing') {
              spreadsheet.updateCell({ formula: `=SUM(M${firstDataRow}:M${lastDataRow})` }, `M${row}`)
              spreadsheet.cellFormat({ color: '#FF0000' }, `M${row}`)
            }
            return
          }

          if (itemType === 'demo_extra_sqft' || itemType === 'demo_extra_rog_sqft') {
            spreadsheet.updateCell({ formula: `=C${row}` }, `J${row}`)
            spreadsheet.updateCell({ formula: `=J${row}*H${row}/27` }, `L${row}`)
            spreadsheet.cellFormat({ color: getColumnBColor(row, parsedData) }, `B${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `L${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `M${row}`)
            return
          }

          if (itemType === 'demo_extra_ft' || itemType === 'demo_extra_rw') {
            spreadsheet.updateCell({ formula: `=C${row}*G${row}` }, `J${row}`)
            spreadsheet.updateCell({ formula: `=J${row}*H${row}/27` }, `L${row}`)
            spreadsheet.cellFormat({ color: getColumnBColor(row, parsedData) }, `B${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `L${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `M${row}`)
            return
          }

          if (itemType === 'demo_extra_ea') {
            spreadsheet.updateCell({ formula: `=C${row}*F${row}*G${row}` }, `J${row}`)
            spreadsheet.updateCell({ formula: `=J${row}*H${row}/27` }, `L${row}`)
            spreadsheet.updateCell({ formula: `=C${row}` }, `M${row}`)
            spreadsheet.cellFormat({ color: getColumnBColor(row, parsedData) }, `B${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `L${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `M${row}`)
            return
          }

          formulas = generateDemolitionFormulas(itemType === 'demolition_item' ? subsection : itemType, row, parsedData)
          // Apply conditional color to column B for demolition items
          if (formulas) {
            spreadsheet.cellFormat({ color: getColumnBColor(row, parsedData) }, `B${row}`)
          }
        } else if (section === 'excavation') {
          if (itemType === 'excavation_sum') {
            const { firstDataRow, lastDataRow } = formulaInfo
            spreadsheet.updateCell({ formula: `=SUM(J${firstDataRow}:J${lastDataRow})` }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `J${row}`)
            spreadsheet.updateCell({ formula: `=SUM(L${firstDataRow}:L${lastDataRow})` }, `L${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `L${row}`)
            return
          }

          if (itemType === 'excavation_havg') {
            const { sumRowNumber } = formulaInfo
            spreadsheet.updateCell({ formula: `=(L${sumRowNumber}*27)/J${sumRowNumber}` }, `C${row}`)
            spreadsheet.cellFormat({ color: getColumnBColor(row, parsedData), fontWeight: 'normal', fontStyle: 'normal' }, `B${row}`)
            spreadsheet.cellFormat({ color: '#000000', fontWeight: 'normal' }, `C${row}`)
            return
          }

          if (itemType === 'backfill_sum') {
            const { firstDataRow, lastDataRow } = formulaInfo
            spreadsheet.updateCell({ formula: `=SUM(J${firstDataRow}:J${lastDataRow})` }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `J${row}`)
            spreadsheet.updateCell({ formula: `=SUM(L${firstDataRow}:L${lastDataRow})` }, `L${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `L${row}`)
            return
          }

          if (itemType === 'mud_slab_sum') {
            const { firstDataRow, lastDataRow } = formulaInfo
            spreadsheet.updateCell({ formula: `=SUM(J${firstDataRow}:J${lastDataRow})` }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `J${row}`)
            spreadsheet.updateCell({ formula: `=SUM(L${firstDataRow}:L${lastDataRow})` }, `L${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `L${row}`)
            return
          }

          if (itemType === 'backfill_extra_sqft') {
            spreadsheet.updateCell({ formula: `=C${row}` }, `J${row}`)
            spreadsheet.updateCell({ formula: `=J${row}*H${row}/27` }, `L${row}`)
            spreadsheet.cellFormat({ color: getColumnBColor(row, parsedData) }, `B${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `L${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `M${row}`)
            return
          }

          if (itemType === 'backfill_extra_ft') {
            spreadsheet.updateCell({ formula: `=C${row}*G${row}` }, `J${row}`)
            spreadsheet.updateCell({ formula: `=J${row}*H${row}/27` }, `L${row}`)
            spreadsheet.cellFormat({ color: getColumnBColor(row, parsedData) }, `B${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `L${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `M${row}`)
            return
          }

          if (itemType === 'backfill_extra_ea') {
            spreadsheet.updateCell({ formula: `=C${row}*F${row}*G${row}` }, `J${row}`)
            spreadsheet.updateCell({ formula: `=J${row}*H${row}/27` }, `L${row}`)
            spreadsheet.updateCell({ formula: `=C${row}` }, `M${row}`)
            spreadsheet.cellFormat({ color: getColumnBColor(row, parsedData) }, `B${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `L${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `M${row}`)
            return
          }

          if (itemType === 'soil_exc_extra_sqft') {
            spreadsheet.updateCell({ formula: `=C${row}` }, `J${row}`)
            spreadsheet.updateCell({ formula: `=J${row}*H${row}/27` }, `K${row}`)
            spreadsheet.updateCell({ formula: `=K${row}*1.3` }, `L${row}`)
            spreadsheet.cellFormat({ color: getColumnBColor(row, parsedData) }, `B${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000', textDecoration: 'line-through' }, `K${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `L${row}`)
            return
          }

          if (itemType === 'soil_exc_extra_ft') {
            spreadsheet.updateCell({ formula: `=C${row}*G${row}` }, `J${row}`)
            spreadsheet.updateCell({ formula: `=J${row}*H${row}/27` }, `K${row}`)
            spreadsheet.updateCell({ formula: `=K${row}*1.3` }, `L${row}`)
            spreadsheet.cellFormat({ color: getColumnBColor(row, parsedData) }, `B${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000', textDecoration: 'line-through' }, `K${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `L${row}`)
            return
          }

          if (itemType === 'soil_exc_extra_ea') {
            spreadsheet.updateCell({ formula: `=C${row}*F${row}*G${row}` }, `J${row}`)
            spreadsheet.updateCell({ formula: `=J${row}*H${row}/27` }, `K${row}`)
            spreadsheet.updateCell({ formula: `=K${row}*1.3` }, `L${row}`)
            spreadsheet.cellFormat({ color: getColumnBColor(row, parsedData) }, `B${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000', textDecoration: 'line-through' }, `K${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `L${row}`)
            return
          }

          formulas = generateExcavationFormulas(itemType === 'excavation_item' ? (parsedData?.itemType || itemType) : itemType, row, parsedData)
          // Apply conditional color to column B for excavation items
          if (formulas) {
            spreadsheet.cellFormat({ color: getColumnBColor(row, parsedData) }, `B${row}`)
          }
        } else if (section === 'rock_excavation') {
          if (itemType === 'rock_excavation_sum') {
            const { firstDataRow, lastDataRow } = formulaInfo
            spreadsheet.updateCell({ formula: `=SUM(J${firstDataRow}:J${lastDataRow})` }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `J${row}`)
            spreadsheet.updateCell({ formula: `=SUM(L${firstDataRow}:L${lastDataRow})` }, `L${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `L${row}`)
            return
          }

          if (itemType === 'rock_excavation_havg') {
            const { sumRowNumber } = formulaInfo
            spreadsheet.updateCell({ formula: `=(L${sumRowNumber}*27)/J${sumRowNumber}` }, `C${row}`)
            spreadsheet.cellFormat({ color: getColumnBColor(row, parsedData), fontWeight: 'normal', fontStyle: 'normal' }, `B${row}`)
            spreadsheet.cellFormat({ color: '#000000', fontWeight: 'normal' }, `C${row}`)
            return
          }

          if (itemType === 'line_drill_sub_header') {
            spreadsheet.cellFormat({ color: '#000000', fontStyle: 'italic', fontWeight: 'normal' }, `E${row}`)
            spreadsheet.cellFormat({ color: '#000000', fontStyle: 'italic', fontWeight: 'normal' }, `H${row}`)
            return
          }

          if (itemType === 'line_drill_concrete_pier') {
            const { refRow } = formulaInfo
            if (refRow) {
              spreadsheet.updateCell({ formula: `=B${refRow}` }, `B${row}`)
              spreadsheet.updateCell({ formula: `=((G${refRow}+F${refRow})*2)*C${refRow}` }, `C${row}`)
              spreadsheet.updateCell({ formula: `=H${refRow}` }, `H${row}`)
            }
            spreadsheet.updateCell({ value: 'FT' }, `D${row}`)
            spreadsheet.updateCell({ formula: `=ROUNDUP(H${row}/2,0)` }, `E${row}`)
            spreadsheet.updateCell({ formula: `=E${row}*C${row}` }, `I${row}`)
            spreadsheet.cellFormat({ color: getColumnBColor(row, parsedData) }, `B${row}`)
            return
          }

          if (itemType === 'line_drill_sewage_pit') {
            const { refRow } = formulaInfo
            if (refRow) {
              spreadsheet.updateCell({ formula: `=B${refRow}` }, `B${row}`)
              spreadsheet.updateCell({ formula: `=SQRT(C${refRow})*4` }, `C${row}`)
              spreadsheet.updateCell({ formula: `=H${refRow}` }, `H${row}`)
            }
            spreadsheet.updateCell({ value: 'FT' }, `D${row}`)
            spreadsheet.updateCell({ formula: `=ROUNDUP(H${row}/2,0)` }, `E${row}`)
            spreadsheet.updateCell({ formula: `=E${row}*C${row}` }, `I${row}`)
            spreadsheet.cellFormat({ color: getColumnBColor(row, parsedData) }, `B${row}`)
            return
          }

          if (itemType === 'line_drill_sump_pit') {
            const { refRow } = formulaInfo
            if (refRow) {
              spreadsheet.updateCell({ formula: `=B${refRow}` }, `B${row}`)
              spreadsheet.updateCell({ formula: `=C${refRow}*8` }, `C${row}`)
            }
            spreadsheet.updateCell({ value: 'FT' }, `D${row}`)
            spreadsheet.updateCell({ formula: `=C${row}` }, `I${row}`)
            spreadsheet.cellFormat({ color: getColumnBColor(row, parsedData) }, `B${row}`)
            return
          }

          if (itemType === 'line_drilling') {
            const rockFormulas = generateRockExcavationFormulas(itemType, row, parsedData)
            if (rockFormulas.qty) spreadsheet.updateCell({ formula: `=${rockFormulas.qty}` }, `E${row}`)
            if (rockFormulas.ft) spreadsheet.updateCell({ formula: `=${rockFormulas.ft}` }, `I${row}`)
            spreadsheet.cellFormat({ color: getColumnBColor(row, parsedData) }, `B${row}`)
            return
          }

          if (itemType === 'line_drill_sum') {
            const { firstDataRow, lastDataRow } = formulaInfo
            spreadsheet.updateCell({ formula: `=SUM(I${firstDataRow}:I${lastDataRow})*2` }, `I${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `I${row}`)
            return
          }

          if (itemType === 'rock_exc_extra_sqft') {
            spreadsheet.updateCell({ formula: `=C${row}` }, `J${row}`)
            spreadsheet.updateCell({ formula: `=J${row}*H${row}/27` }, `L${row}`)
            spreadsheet.cellFormat({ color: getColumnBColor(row, parsedData) }, `B${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `L${row}`)
            return
          }

          if (itemType === 'rock_exc_extra_ft') {
            spreadsheet.updateCell({ formula: `=C${row}*G${row}` }, `J${row}`)
            spreadsheet.updateCell({ formula: `=J${row}*H${row}/27` }, `L${row}`)
            spreadsheet.cellFormat({ color: getColumnBColor(row, parsedData) }, `B${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `L${row}`)
            return
          }

          if (itemType === 'rock_exc_extra_ea') {
            spreadsheet.updateCell({ formula: `=C${row}*F${row}*G${row}` }, `J${row}`)
            spreadsheet.updateCell({ formula: `=J${row}*H${row}/27` }, `L${row}`)
            spreadsheet.updateCell({ formula: `=C${row}` }, `M${row}`)
            spreadsheet.cellFormat({ color: getColumnBColor(row, parsedData) }, `B${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `L${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `M${row}`)
            return
          }

          formulas = generateRockExcavationFormulas(itemType === 'rock_excavation_item' ? (parsedData?.itemType || itemType) : itemType, row, parsedData)
        } else if (section === 'soe') {
          // SOE section - use generateSoeFormulas for items
          const soeItemTypes = ['soldier_pile_item', 'timber_brace_item', 'timber_waler_item', 'timber_stringer_item', 'drilled_hole_grout_item', 'soe_generic_item', 'backpacking_item', 'supporting_angle', 'parging', 'heel_block', 'underpinning', 'shims', 'rock_anchor', 'rock_bolt', 'anchor', 'tie_back', 'concrete_soil_retention_pier', 'guide_wall', 'dowel_bar', 'rock_pin', 'shotcrete', 'permission_grouting', 'button', 'rock_stabilization', 'form_board']

          if (soeItemTypes.includes(itemType)) {
            const soeFormulas = generateSoeFormulas(itemType, row, parsedData || formulaInfo)
            if (soeFormulas.takeoff) spreadsheet.updateCell({ formula: `=${soeFormulas.takeoff}` }, `C${row}`)
            if (soeFormulas.length !== undefined) {
              if (typeof soeFormulas.length === 'string') {
                spreadsheet.updateCell({ formula: `=${soeFormulas.length}` }, `F${row}`)
              } else {
                spreadsheet.updateCell({ value: soeFormulas.length }, `F${row}`)
              }
            }
            if (soeFormulas.width !== undefined) {
              if (typeof soeFormulas.width === 'string') {
                spreadsheet.updateCell({ formula: `=${soeFormulas.width}` }, `G${row}`)
              } else {
                spreadsheet.updateCell({ value: soeFormulas.width }, `G${row}`)
              }
            }
            if (soeFormulas.height !== undefined) {
              if (typeof soeFormulas.height === 'string') {
                spreadsheet.updateCell({ formula: `=${soeFormulas.height}` }, `H${row}`)
              } else {
                spreadsheet.updateCell({ value: soeFormulas.height }, `H${row}`)
              }
            }
            if (soeFormulas.qty !== undefined) {
              if (typeof soeFormulas.qty === 'string') {
                spreadsheet.updateCell({ formula: `=${soeFormulas.qty}` }, `E${row}`)
              } else {
                spreadsheet.updateCell({ value: soeFormulas.qty }, `E${row}`)
              }
            }
            if (soeFormulas.ft) {
              spreadsheet.updateCell({ formula: `=${soeFormulas.ft}` }, `I${row}`)
            }
            if (soeFormulas.sqFt) {
              spreadsheet.updateCell({ formula: `=${soeFormulas.sqFt}` }, `J${row}`)
            }
            if (soeFormulas.lbs) {
              spreadsheet.updateCell({ formula: `=${soeFormulas.lbs}` }, `K${row}`)
            }
            if (soeFormulas.cy) {
              spreadsheet.updateCell({ formula: `=${soeFormulas.cy}` }, `L${row}`)
            }
            if (soeFormulas.qtyFinal) {
              spreadsheet.updateCell({ formula: `=${soeFormulas.qtyFinal}` }, `M${row}`)
            }
            spreadsheet.cellFormat({ color: getColumnBColor(row, parsedData) }, `B${row}`)

            if (itemType === 'backpacking_item') {
              spreadsheet.cellFormat({ color: '#000000' }, `C${row}`)
              spreadsheet.cellFormat({ color: '#FF0000' }, `J${row}`)
            }
            if (itemType === 'shims') {
              spreadsheet.cellFormat({ color: '#FF0000' }, `I${row}`)
              spreadsheet.cellFormat({ color: '#FF0000' }, `J${row}`)
            }
            if (itemType === 'timber_brace_item' && !formulaInfo?.hasMultipleItems) {
              spreadsheet.cellFormat({ color: '#FF0000' }, `I${row}`)
              spreadsheet.cellFormat({ color: '#FF0000' }, `M${row}`)
            }
            if (itemType === 'timber_waler_item' && !formulaInfo?.hasMultipleItems) {
              spreadsheet.cellFormat({ color: '#FF0000' }, `I${row}`)
              spreadsheet.cellFormat({ color: '#FF0000' }, `M${row}`)
            }
            if (itemType === 'timber_stringer_item' && !formulaInfo?.hasMultipleItems) {
              spreadsheet.cellFormat({ color: '#FF0000' }, `I${row}`)
            }
            if (itemType === 'drilled_hole_grout_item' && !formulaInfo?.hasMultipleItems) {
              spreadsheet.cellFormat({ color: '#FF0000' }, `I${row}`)
              spreadsheet.cellFormat({ color: '#FF0000' }, `J${row}`)
              spreadsheet.cellFormat({ color: '#FF0000' }, `L${row}`)
              spreadsheet.cellFormat({ color: '#FF0000' }, `M${row}`)
            }
            return
          }

          if (itemType === 'soldier_pile_group_sum' || itemType === 'timber_soldier_pile_group_sum' || itemType === 'timber_plank_group_sum' || itemType === 'timber_raker_group_sum' || itemType === 'timber_brace_group_sum' || itemType === 'timber_waler_group_sum' || itemType === 'timber_stringer_group_sum' || itemType === 'timber_post_group_sum' || itemType === 'vertical_timber_sheets_group_sum' || itemType === 'horizontal_timber_sheets_group_sum' || itemType === 'drilled_hole_grout_group_sum' || itemType === 'soe_generic_sum') {
            const { firstDataRow, lastDataRow, subsectionName } = formulaInfo
            const ftSumSubsections = ['Rock anchors', 'Rock bolts', 'Anchor', 'Tie back', 'Dowel bar', 'Rock pins', 'Shotcrete', 'Permission grouting', 'Form board', 'Guide wall']
            const sqFtSubsections = ['Sheet pile', 'Timber lagging', 'Timber sheeting', 'Vertical timber sheets', 'Horizontal timber sheets', 'Parging', 'Heel blocks', 'Underpinning', 'Concrete soil retention piers', 'Guide wall', 'Shotcrete', 'Permission grouting', 'Buttons', 'Form board', 'Rock stabilization']
            const lbsSubsections = ['Primary secant piles', 'Secondary secant piles', 'Tangent piles', 'Sheet pile', 'Waler', 'Raker', 'Upper Raker', 'Lower Raker', 'Stand off', 'Kicker', 'Channel', 'Roll chock', 'Stud beam', 'Inner corner brace', 'Knee brace', 'Supporting angle']
            const qtySubsections = ['Primary secant piles', 'Secondary secant piles', 'Tangent piles', 'Waler', 'Raker', 'Upper Raker', 'Lower Raker', 'Stand off', 'Kicker', 'Channel', 'Roll chock', 'Stud beam', 'Inner corner brace', 'Knee brace', 'Supporting angle', 'Heel blocks', 'Underpinning', 'Rock anchors', 'Rock bolts', 'Anchor', 'Tie back', 'Concrete soil retention piers', 'Dowel bar', 'Rock pins', 'Buttons']
            const cySubsections = ['Heel blocks', 'Underpinning', 'Concrete soil retention piers', 'Guide wall', 'Shotcrete', 'Buttons', 'Rock stabilization']

            if ((itemType === 'timber_waler_group_sum' || itemType === 'timber_brace_group_sum' || itemType === 'timber_stringer_group_sum' || itemType === 'drilled_hole_grout_group_sum') || (itemType !== 'timber_brace_group_sum' && itemType !== 'timber_stringer_group_sum' && itemType !== 'drilled_hole_grout_group_sum' && subsectionName !== 'Heel blocks' && (ftSumSubsections.includes(subsectionName) || !['Concrete soil retention piers', 'Buttons', 'Rock stabilization'].includes(subsectionName)))) {
              spreadsheet.updateCell({ formula: `=SUM(I${firstDataRow}:I${lastDataRow})` }, `I${row}`)
              spreadsheet.cellFormat({ color: '#FF0000' }, `I${row}`)
            }
            if (sqFtSubsections.includes(subsectionName) || itemType === 'drilled_hole_grout_group_sum') {
              spreadsheet.updateCell({ formula: `=SUM(J${firstDataRow}:J${lastDataRow})` }, `J${row}`)
              spreadsheet.cellFormat({ color: '#FF0000' }, `J${row}`)
            }
            if (itemType === 'soldier_pile_group_sum' || lbsSubsections.includes(subsectionName)) {
              spreadsheet.updateCell({ formula: `=SUM(K${firstDataRow}:K${lastDataRow})` }, `K${row}`)
              spreadsheet.cellFormat({ color: '#FF0000' }, `K${row}`)
            }
            if ((itemType === 'soldier_pile_group_sum' || itemType === 'timber_soldier_pile_group_sum' || itemType === 'timber_plank_group_sum' || itemType === 'timber_raker_group_sum' || itemType === 'timber_waler_group_sum' || itemType === 'timber_brace_group_sum' || itemType === 'timber_post_group_sum' || itemType === 'drilled_hole_grout_group_sum') || qtySubsections.includes(subsectionName)) {
              spreadsheet.updateCell({ formula: `=SUM(M${firstDataRow}:M${lastDataRow})` }, `M${row}`)
              spreadsheet.cellFormat({ color: '#FF0000' }, `M${row}`)
            }
            if (cySubsections.includes(subsectionName) || itemType === 'drilled_hole_grout_group_sum') {
              spreadsheet.updateCell({ formula: `=SUM(L${firstDataRow}:L${lastDataRow})` }, `L${row}`)
              spreadsheet.cellFormat({ color: '#FF0000' }, `L${row}`)
            }
            return
          }
        } else if (section === 'foundation') {
          // Foundation section continues in the full Spreadsheet.jsx - copy all the logic
          if (itemType === 'foundation_piles_misc') {
            spreadsheet.cellFormat({ color: '#FF0000' }, `B${row}`)
            return
          }
          if (itemType === 'foundation_sum') {
            const { firstDataRow, lastDataRow, subsectionName, isDualDiameter, excludeISum, excludeJSum, cySumOnly, lSumRange } = formulaInfo

            if (!excludeISum) {
              const ftSumSubsections = ['Helical foundation pile', 'Driven foundation pile', 'Stelcor drilled displacement pile', 'CFA pile', 'Grade beams', 'Tie beam', 'Thickened slab', 'Corbel', 'Linear Wall', 'Foundation Wall', 'Retaining walls', 'Barrier wall', 'Drilled foundation pile', 'Strip Footings', 'Stem wall', 'Elevator Pit', 'Service elevator pit', 'Detention tank', 'Duplex sewage ejector pit', 'Deep sewage ejector pit', 'Sump pump pit', 'Grease trap', 'House trap', 'SOG', 'Stairs on grade Stairs', 'Electric conduit']
              if (ftSumSubsections.includes(subsectionName)) {
                spreadsheet.updateCell({ formula: `=SUM(I${firstDataRow}:I${lastDataRow})` }, `I${row}`)
                spreadsheet.cellFormat({ color: '#FF0000' }, `I${row}`)
              }
            }

            if (subsectionName === 'Drilled foundation pile' && isDualDiameter) {
              spreadsheet.updateCell({ formula: `=SUM(I${firstDataRow}:I${lastDataRow})` }, `I${row}`)
              spreadsheet.updateCell({ formula: `=SUM(J${firstDataRow}:J${lastDataRow})` }, `J${row}`)
              spreadsheet.cellFormat({ color: '#FF0000' }, `I${row}`)
              spreadsheet.cellFormat({ color: '#FF0000' }, `J${row}`)
            }

            if (!excludeJSum && !cySumOnly) {
              const sqFtSubsections = ['Pile caps', 'Isolated Footings', 'Pilaster', 'Pier', 'Strip Footings', 'Grade beams', 'Tie beam', 'Thickened slab', 'Corbel', 'Linear Wall', 'Foundation Wall', 'Retaining walls', 'Barrier wall', 'Stem wall', 'Elevator Pit', 'Service elevator pit', 'Detention tank', 'Duplex sewage ejector pit', 'Deep sewage ejector pit', 'Sump pump pit', 'Grease trap', 'House trap', 'Mat slab', 'SOG', 'Stairs on grade Stairs']
              if (sqFtSubsections.includes(subsectionName)) {
                spreadsheet.updateCell({ formula: `=SUM(J${firstDataRow}:J${lastDataRow})` }, `J${row}`)
                spreadsheet.cellFormat({ color: '#FF0000' }, `J${row}`)
              }
            }

            const lbsSubsections = ['Drilled foundation pile', 'Helical foundation pile', 'Driven foundation pile', 'Stelcor drilled displacement pile']
            if (lbsSubsections.includes(subsectionName)) {
              spreadsheet.updateCell({ formula: `=SUM(K${firstDataRow}:K${lastDataRow})` }, `K${row}`)
              spreadsheet.cellFormat({ color: '#FF0000' }, `K${row}`)
            }

            const qtySubsections = ['Drilled foundation pile', 'Helical foundation pile', 'Driven foundation pile', 'Stelcor drilled displacement pile', 'CFA pile', 'Pile caps', 'Isolated Footings', 'Pilaster', 'Pier', 'Stairs on grade Stairs']
            if (qtySubsections.includes(subsectionName)) {
              spreadsheet.updateCell({ formula: `=SUM(M${firstDataRow}:M${lastDataRow})` }, `M${row}`)
              spreadsheet.cellFormat({ color: '#FF0000' }, `M${row}`)
            }

            // CY sum
            if (cySumOnly) {
              spreadsheet.updateCell({ formula: `=SUM(L${firstDataRow}:L${lastDataRow})` }, `L${row}`)
              spreadsheet.cellFormat({ color: '#FF0000' }, `L${row}`)
            } else if (!formulaInfo.excludeLSum) {
              const cySubsections = ['Pile caps', 'Strip Footings', 'Isolated Footings', 'Pilaster', 'Grade beams', 'Tie beam', 'Thickened slab', 'Pier', 'Corbel', 'Linear Wall', 'Foundation Wall', 'Retaining walls', 'Barrier wall', 'Stem wall', 'Elevator Pit', 'Service elevator pit', 'Detention tank', 'Duplex sewage ejector pit', 'Deep sewage ejector pit', 'Sump pump pit', 'Grease trap', 'House trap', 'Mat slab', 'SOG', 'Stairs on grade Stairs']
              if (cySubsections.includes(subsectionName)) {
                const lFormula = lSumRange ? `=SUM(${lSumRange})` : `=SUM(L${firstDataRow}:L${lastDataRow})`
                deferredFoundationSumL.push({ row, lFormula })
                spreadsheet.cellFormat({ color: '#FF0000' }, `L${row}`)
              }
            }
            return
          }

          if (itemType === 'foundation_section_cy_sum') {
            const { sumRows } = formulaInfo
            if (sumRows && sumRows.length > 0) {
              const sumRefs = sumRows.map((r) => `L${r}`).join(',')
              spreadsheet.updateCell({ formula: `=SUM(${sumRefs})` }, `C${row}`)
            }
            return
          }

          if (itemType === 'foundation_extra_sqft') {
            spreadsheet.updateCell({ formula: `=C${row}` }, `J${row}`)
            spreadsheet.updateCell({ formula: `=J${row}*H${row}/27` }, `L${row}`)
            spreadsheet.cellFormat({ color: getColumnBColor(row, parsedData) }, `B${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `L${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `M${row}`)
            return
          }

          if (itemType === 'foundation_extra_ft') {
            spreadsheet.updateCell({ formula: `=C${row}*G${row}` }, `J${row}`)
            spreadsheet.updateCell({ formula: `=J${row}*H${row}/27` }, `L${row}`)
            spreadsheet.cellFormat({ color: getColumnBColor(row, parsedData) }, `B${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `L${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `M${row}`)
            return
          }

          if (itemType === 'foundation_extra_ea') {
            spreadsheet.updateCell({ formula: `=C${row}*F${row}*G${row}` }, `J${row}`)
            spreadsheet.updateCell({ formula: `=J${row}*H${row}/27` }, `L${row}`)
            spreadsheet.updateCell({ formula: `=C${row}` }, `M${row}`)
            spreadsheet.cellFormat({ color: getColumnBColor(row, parsedData) }, `B${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `L${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `M${row}`)
            return
          }

          // Special handling for Sump pit row (red color for J, L, M columns)
          if ((itemType === 'elevator_pit' || itemType === 'service_elevator_pit') && parsedData?.parsed?.itemSubType === 'sump_pit') {
            const foundationFormulas = generateFoundationFormulas(itemType, row, parsedData || formulaInfo)
            if (foundationFormulas.sqFt) spreadsheet.updateCell({ formula: `=${foundationFormulas.sqFt}` }, `J${row}`)
            if (foundationFormulas.cy) spreadsheet.updateCell({ formula: `=${foundationFormulas.cy}` }, `L${row}`)
            if (foundationFormulas.qtyFinal) spreadsheet.updateCell({ formula: `=${foundationFormulas.qtyFinal}` }, `M${row}`)
            // Apply red color to J, L, M for Sump pit
            spreadsheet.cellFormat({ color: '#FF0000' }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `L${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `M${row}`)
            return
          }

          // Foundation items using generateFoundationFormulas
          const foundationItemTypes = ['drilled_foundation_pile', 'helical_foundation_pile', 'driven_foundation_pile', 'stelcor_drilled_displacement_pile', 'cfa_pile', 'pile_cap', 'strip_footing', 'isolated_footing', 'pilaster', 'grade_beam', 'tie_beam', 'thickened_slab', 'buttress_takeoff', 'buttress_final', 'pier', 'corbel', 'linear_wall', 'foundation_wall', 'retaining_wall', 'barrier_wall', 'stem_wall', 'elevator_pit', 'service_elevator_pit', 'detention_tank', 'duplex_sewage_ejector_pit', 'deep_sewage_ejector_pit', 'sump_pump_pit', 'grease_trap', 'house_trap', 'mat_slab', 'mud_slab_foundation', 'sog', 'stairs_on_grade', 'electric_conduit']
          if (foundationItemTypes.includes(itemType)) {
            const foundationFormulas = generateFoundationFormulas(itemType, row, parsedData || formulaInfo)
            if (foundationFormulas.takeoff) spreadsheet.updateCell({ formula: `=${foundationFormulas.takeoff}` }, `C${row}`)
            if (foundationFormulas.length != null) {
              if (typeof foundationFormulas.length === 'string') {
                spreadsheet.updateCell({ formula: `=${foundationFormulas.length}` }, `F${row}`)
              } else {
                spreadsheet.updateCell({ value: foundationFormulas.length }, `F${row}`)
              }
            }
            if (foundationFormulas.width != null) {
              if (typeof foundationFormulas.width === 'string') {
                spreadsheet.updateCell({ formula: `=${foundationFormulas.width}` }, `G${row}`)
              } else {
                spreadsheet.updateCell({ value: foundationFormulas.width }, `G${row}`)
              }
            }
            if (foundationFormulas.height != null) {
              if (typeof foundationFormulas.height === 'string') {
                spreadsheet.updateCell({ formula: `=${foundationFormulas.height}` }, `H${row}`)
              } else {
                spreadsheet.updateCell({ value: foundationFormulas.height }, `H${row}`)
              }
            }
            if (foundationFormulas.qty !== undefined) {
              if (typeof foundationFormulas.qty === 'string') {
                spreadsheet.updateCell({ formula: `=${foundationFormulas.qty}` }, `E${row}`)
              } else {
                spreadsheet.updateCell({ value: foundationFormulas.qty }, `E${row}`)
              }
            }
            if (foundationFormulas.ft) spreadsheet.updateCell({ formula: `=${foundationFormulas.ft}` }, `I${row}`)
            if (foundationFormulas.sqFt) spreadsheet.updateCell({ formula: `=${foundationFormulas.sqFt}` }, `J${row}`)
            if (foundationFormulas.sqFt2) spreadsheet.updateCell({ formula: `=${foundationFormulas.sqFt2}` }, `J${row}`)
            if (foundationFormulas.lbs) spreadsheet.updateCell({ formula: `=${foundationFormulas.lbs}` }, `K${row}`)
            if (foundationFormulas.cy) spreadsheet.updateCell({ formula: `=${foundationFormulas.cy}` }, `L${row}`)
            if (foundationFormulas.qtyFinal) spreadsheet.updateCell({ formula: `=${foundationFormulas.qtyFinal}` }, `M${row}`)
            spreadsheet.cellFormat({ color: getColumnBColor(row, parsedData) }, `B${row}`)
            return
          }
        } else if (section === 'waterproofing') {
          // Waterproofing formulas
          if (itemType === 'waterproofing_exterior_side_header') {
            spreadsheet.cellFormat(
              { color: '#000000', fontStyle: 'italic', fontWeight: 'bold', textDecoration: 'underline', backgroundColor: '#D9E1F2' },
              `H${row}`
            )
            return
          }
          if (itemType === 'waterproofing_exterior_side_sum') {
            const { firstDataRow, lastDataRow } = formulaInfo
            spreadsheet.updateCell({ formula: `=SUM(I${firstDataRow}:I${lastDataRow})` }, `I${row}`)
            spreadsheet.updateCell({ formula: `=SUM(J${firstDataRow}:J${lastDataRow})` }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `I${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `J${row}`)
            return
          }

          if (itemType === 'waterproofing_negative_side_sum') {
            const { firstDataRow, lastDataRow } = formulaInfo
            spreadsheet.updateCell({ formula: `=SUM(I${firstDataRow}:I${lastDataRow})` }, `I${row}`)
            spreadsheet.updateCell({ formula: `=SUM(J${firstDataRow}:J${lastDataRow})` }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `I${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `J${row}`)
            return
          }

          if (itemType === 'waterproofing_horizontal_wp') {
            spreadsheet.updateCell({ formula: `=C${row}` }, `J${row}`)
            spreadsheet.cellFormat({ color: getColumnBColor(row, parsedData) }, `B${row}`)
            return
          }

          if (itemType === 'waterproofing_horizontal_wp_sum') {
            const { firstDataRow, lastDataRow } = formulaInfo
            spreadsheet.updateCell({ formula: `=SUM(J${firstDataRow}:J${lastDataRow})` }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `J${row}`)
            return
          }

          if (itemType === 'waterproofing_horizontal_insulation') {
            spreadsheet.updateCell({ formula: `=C${row}` }, `J${row}`)
            spreadsheet.cellFormat({ color: getColumnBColor(row, parsedData) }, `B${row}`)
            return
          }

          if (itemType === 'waterproofing_horizontal_insulation_sum') {
            const { firstDataRow, lastDataRow } = formulaInfo
            spreadsheet.updateCell({ formula: `=SUM(J${firstDataRow}:J${lastDataRow})` }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `J${row}`)
            return
          }

          const waterproofingFormulas = generateWaterproofingFormulas(itemType, row, parsedData || formulaInfo)
          if (waterproofingFormulas.ft) spreadsheet.updateCell({ formula: `=${waterproofingFormulas.ft}` }, `I${row}`)
          if (waterproofingFormulas.height != null) spreadsheet.updateCell({ value: waterproofingFormulas.height }, `H${row}`)
          if (waterproofingFormulas.heightFormula) spreadsheet.updateCell({ formula: `=${waterproofingFormulas.heightFormula}` }, `H${row}`)
          if (waterproofingFormulas.sqFt) spreadsheet.updateCell({ formula: `=${waterproofingFormulas.sqFt}` }, `J${row}`)
          if (waterproofingFormulas.cy) spreadsheet.updateCell({ formula: `=${waterproofingFormulas.cy}` }, `L${row}`)
          return
        } else if (section === 'trenching') {
          if (itemType === 'trenching_section_header') {
            const { patchbackRow } = formulaInfo
            if (patchbackRow != null) {
              spreadsheet.updateCell({ formula: `=L${patchbackRow}` }, `C${row}`)
            }
            return
          }
          if (itemType === 'trenching_item') {
            const { takeoffRefRow, hFormula, lYellow } = formulaInfo
            if (takeoffRefRow != null) {
              spreadsheet.updateCell({ formula: `=C${takeoffRefRow}` }, `C${row}`)
            }
            if (hFormula) {
              spreadsheet.updateCell({ formula: `=${hFormula}` }, `H${row}`)
            }
            spreadsheet.updateCell({ formula: `=C${row}` }, `I${row}`)
            spreadsheet.updateCell({ formula: `=I${row}*G${row}` }, `J${row}`)
            spreadsheet.updateCell({ formula: `=J${row}*H${row}/27` }, `L${row}`)
            spreadsheet.cellFormat({ color: getColumnBColor(row, parsedData) }, `B${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `I${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `L${row}`)
            if (lYellow) {
              spreadsheet.cellFormat({ backgroundColor: '#FFF2CC' }, `L${row}`)
            }
            return
          }
        } else if (section === 'superstructure') {
          // Superstructure section - complete formula logic from Spreadsheet.jsx
          if (itemType === 'superstructure_sum') {
            const { firstDataRow, lastDataRow } = formulaInfo
            spreadsheet.updateCell({ formula: `=SUM(J${firstDataRow}:J${lastDataRow})` }, `J${row}`)
            spreadsheet.updateCell({ formula: `=SUM(L${firstDataRow}:L${lastDataRow})` }, `L${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `L${row}`)
            return
          }
          if (itemType === 'superstructure_slab_steps_sum') {
            const { firstDataRow, lastDataRow } = formulaInfo
            spreadsheet.updateCell({ formula: `=SUM(I${firstDataRow}:I${lastDataRow})` }, `I${row}`)
            spreadsheet.updateCell({ formula: `=SUM(J${firstDataRow}:J${lastDataRow})` }, `J${row}`)
            spreadsheet.updateCell({ formula: `=SUM(L${firstDataRow}:L${lastDataRow})` }, `L${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `I${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `L${row}`)
            return
          }
          if (itemType === 'superstructure_lw_concrete_fill_sum') {
            const { firstDataRow, lastDataRow } = formulaInfo
            spreadsheet.updateCell({ formula: `=SUM(J${firstDataRow}:J${lastDataRow})` }, `J${row}`)
            spreadsheet.updateCell({ formula: `=SUM(L${firstDataRow}:L${lastDataRow})` }, `L${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `L${row}`)
            return
          }
          if (itemType === 'superstructure_item') {
            const parsed = parsedData || formulaInfo
            const heightFormula = parsed?.parsed?.heightFormula
            const heightValue = parsed?.parsed?.heightValue
            if (heightFormula) spreadsheet.updateCell({ formula: `=${heightFormula}` }, `H${row}`)
            if (heightValue != null) spreadsheet.updateCell({ value: heightValue }, `H${row}`)
            spreadsheet.updateCell({ formula: `=C${row}` }, `J${row}`)
            spreadsheet.updateCell({ formula: `=J${row}*H${row}/27` }, `L${row}`)
            spreadsheet.cellFormat({ color: getColumnBColor(row, parsedData) }, `B${row}`)
            return
          }
          if (itemType === 'superstructure_slab_step') {
            spreadsheet.updateCell({ formula: `=C${row}` }, `I${row}`)
            spreadsheet.updateCell({ formula: `=I${row}*H${row}` }, `J${row}`)
            spreadsheet.updateCell({ formula: `=J${row}*G${row}/27` }, `L${row}`)
            spreadsheet.cellFormat({ color: getColumnBColor(row, parsedData) }, `B${row}`)
            return
          }
          if (itemType === 'superstructure_lw_concrete_fill') {
            spreadsheet.updateCell({ formula: `=C${row}` }, `J${row}`)
            spreadsheet.updateCell({ formula: `=J${row}*H${row}/27` }, `L${row}`)
            spreadsheet.cellFormat({ color: getColumnBColor(row, parsedData) }, `B${row}`)
            return
          }
          if (itemType === 'superstructure_somd_item') {
            spreadsheet.cellFormat({ color: getColumnBColor(row, parsedData) }, `B${row}`)
            return
          }
          if (itemType === 'superstructure_somd_gen1') {
            const { firstDataRow, lastDataRow, heightFormula } = formulaInfo
            spreadsheet.updateCell({ formula: `=SUM(C${firstDataRow}:C${lastDataRow})` }, `C${row}`)
            if (heightFormula) spreadsheet.updateCell({ formula: `=${heightFormula}` }, `H${row}`)
            spreadsheet.updateCell({ formula: `=C${row}` }, `J${row}`)
            spreadsheet.updateCell({ formula: `=(J${row}*H${row})/27` }, `L${row}`)
            spreadsheet.cellFormat({ color: getColumnBColor(row, parsedData) }, `B${row}`)
            spreadsheet.cellFormat({ backgroundColor: '#DDEBF7' }, `H${row}`)
            return
          }
          if (itemType === 'superstructure_somd_gen2') {
            const { takeoffRefRow, heightValue } = formulaInfo
            if (takeoffRefRow != null) spreadsheet.updateCell({ formula: `=C${takeoffRefRow}` }, `C${row}`)
            if (heightValue != null) spreadsheet.updateCell({ value: heightValue }, `H${row}`)
            spreadsheet.updateCell({ formula: `=C${row}` }, `J${row}`)
            spreadsheet.updateCell({ formula: `=(J${row}*H${row})/27/2` }, `L${row}`)
            spreadsheet.cellFormat({ backgroundColor: '#DDEBF7' }, `H${row}`)
            return
          }
          if (itemType === 'superstructure_somd_sum') {
            const { gen1Row, gen2Row } = formulaInfo
            if (gen1Row != null) spreadsheet.updateCell({ formula: `=J${gen1Row}` }, `J${row}`)
            if (gen1Row != null && gen2Row != null) spreadsheet.updateCell({ formula: `=SUM(L${gen1Row}:L${gen2Row})` }, `L${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `L${row}`)
            return
          }
          if (itemType === 'superstructure_topping_slab') {
            spreadsheet.updateCell({ formula: `=C${row}` }, `J${row}`)
            spreadsheet.updateCell({ formula: `=J${row}*H${row}/27` }, `L${row}`)
            spreadsheet.cellFormat({ color: getColumnBColor(row, parsedData) }, `B${row}`)
            return
          }
          if (itemType === 'superstructure_topping_slab_sum') {
            const { firstDataRow, lastDataRow } = formulaInfo
            spreadsheet.updateCell({ formula: `=SUM(J${firstDataRow}:J${lastDataRow})` }, `J${row}`)
            spreadsheet.updateCell({ formula: `=SUM(L${firstDataRow}:L${lastDataRow})` }, `L${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `L${row}`)
            return
          }
          if (itemType === 'superstructure_thermal_break') {
            const parsed = parsedData || formulaInfo
            const qty = parsed?.parsed?.qty
            if (qty != null) {
              spreadsheet.updateCell({ formula: `=C${row}*E${row}` }, `I${row}`)
            } else {
              spreadsheet.updateCell({ formula: `=C${row}` }, `I${row}`)
            }
            spreadsheet.cellFormat({ color: getColumnBColor(row, parsedData) }, `B${row}`)
            return
          }
          if (itemType === 'superstructure_thermal_break_sum') {
            const { firstDataRow, lastDataRow } = formulaInfo
            spreadsheet.updateCell({ formula: `=SUM(I${firstDataRow}:I${lastDataRow})` }, `I${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `I${row}`)
            return
          }
          if (itemType === 'superstructure_raised_knee_wall') {
            spreadsheet.updateCell({ formula: `=C${row}` }, `I${row}`)
            spreadsheet.updateCell({ formula: `=I${row}*H${row}` }, `J${row}`)
            spreadsheet.updateCell({ formula: `=J${row}*G${row}/27` }, `L${row}`)
            spreadsheet.cellFormat({ color: getColumnBColor(row, parsedData) }, `B${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `I${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `L${row}`)
            return
          }
          if (itemType === 'superstructure_raised_styrofoam') {
            const { takeoffRefRow, heightRefRow } = formulaInfo
            if (takeoffRefRow != null) spreadsheet.updateCell({ formula: `=C${takeoffRefRow}` }, `C${row}`)
            if (heightRefRow != null) spreadsheet.updateCell({ formula: `=H${heightRefRow}` }, `H${row}`)
            spreadsheet.updateCell({ formula: `=C${row}` }, `J${row}`)
            spreadsheet.updateCell({ formula: `=J${row}*H${row}/27` }, `L${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `L${row}`)
            return
          }
          if (itemType === 'superstructure_raised_slab') {
            spreadsheet.updateCell({ formula: `=C${row}` }, `J${row}`)
            spreadsheet.updateCell({ formula: `=J${row}*H${row}/27` }, `L${row}`)
            spreadsheet.cellFormat({ color: getColumnBColor(row, parsedData) }, `B${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `L${row}`)
            return
          }
          if (itemType === 'superstructure_raised_slab_sum') {
            const { firstDataRow, lastDataRow } = formulaInfo
            spreadsheet.updateCell({ formula: `=SUM(I${firstDataRow}:I${lastDataRow})` }, `I${row}`)
            spreadsheet.updateCell({ formula: `=SUM(J${firstDataRow}:J${lastDataRow})` }, `J${row}`)
            spreadsheet.updateCell({ formula: `=SUM(L${firstDataRow}:L${lastDataRow})` }, `L${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `I${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `L${row}`)
            return
          }
          if (itemType === 'superstructure_builtup_knee_wall') {
            spreadsheet.updateCell({ formula: `=C${row}` }, `I${row}`)
            spreadsheet.updateCell({ formula: `=I${row}*H${row}` }, `J${row}`)
            spreadsheet.updateCell({ formula: `=J${row}*G${row}/27` }, `L${row}`)
            spreadsheet.cellFormat({ color: getColumnBColor(row, parsedData) }, `B${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `I${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `L${row}`)
            return
          }
          if (itemType === 'superstructure_builtup_styrofoam') {
            const { takeoffRefRow, heightRefRow } = formulaInfo
            if (takeoffRefRow != null) spreadsheet.updateCell({ formula: `=C${takeoffRefRow}` }, `C${row}`)
            if (heightRefRow != null) spreadsheet.updateCell({ formula: `=H${heightRefRow}` }, `H${row}`)
            spreadsheet.updateCell({ formula: `=C${row}` }, `J${row}`)
            spreadsheet.updateCell({ formula: `=J${row}*H${row}/27` }, `L${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `L${row}`)
            return
          }
          if (itemType === 'superstructure_builtup_slab') {
            spreadsheet.updateCell({ formula: `=C${row}` }, `J${row}`)
            spreadsheet.updateCell({ formula: `=J${row}*H${row}/27` }, `L${row}`)
            spreadsheet.cellFormat({ color: getColumnBColor(row, parsedData) }, `B${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `L${row}`)
            return
          }
          if (itemType === 'superstructure_builtup_slab_sum') {
            const { firstDataRow, lastDataRow } = formulaInfo
            spreadsheet.updateCell({ formula: `=SUM(I${firstDataRow}:I${lastDataRow})` }, `I${row}`)
            spreadsheet.updateCell({ formula: `=SUM(J${firstDataRow}:J${lastDataRow})` }, `J${row}`)
            spreadsheet.updateCell({ formula: `=SUM(L${firstDataRow}:L${lastDataRow})` }, `L${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `I${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `L${row}`)
            return
          }
          if (itemType === 'superstructure_builtup_ramps_knee_wall') {
            spreadsheet.updateCell({ formula: `=C${row}` }, `I${row}`)
            spreadsheet.updateCell({ formula: `=I${row}*H${row}` }, `J${row}`)
            spreadsheet.updateCell({ formula: `=J${row}*G${row}/27` }, `L${row}`)
            spreadsheet.cellFormat({ color: getColumnBColor(row, parsedData) }, `B${row}`)
            return
          }
          if (itemType === 'superstructure_builtup_ramps_knee_sum') {
            const { firstDataRow, lastDataRow } = formulaInfo
            spreadsheet.updateCell({ formula: `=SUM(I${firstDataRow}:I${lastDataRow})` }, `I${row}`)
            spreadsheet.updateCell({ formula: `=SUM(J${firstDataRow}:J${lastDataRow})` }, `J${row}`)
            spreadsheet.updateCell({ formula: `=SUM(L${firstDataRow}:L${lastDataRow})` }, `L${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `I${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `L${row}`)
            return
          }
          if (itemType === 'superstructure_builtup_ramps_styrofoam') {
            const { takeoffRefRow, heightRefRow } = formulaInfo
            if (takeoffRefRow != null) spreadsheet.updateCell({ formula: `=C${takeoffRefRow}` }, `C${row}`)
            if (heightRefRow != null) spreadsheet.updateCell({ formula: `=H${heightRefRow}` }, `H${row}`)
            spreadsheet.updateCell({ formula: `=C${row}` }, `J${row}`)
            spreadsheet.updateCell({ formula: `=J${row}*H${row}/27` }, `L${row}`)
            return
          }
          if (itemType === 'superstructure_builtup_ramps_styro_sum') {
            const { firstDataRow, lastDataRow } = formulaInfo
            spreadsheet.updateCell({ formula: `=SUM(J${firstDataRow}:J${lastDataRow})` }, `J${row}`)
            spreadsheet.updateCell({ formula: `=SUM(L${firstDataRow}:L${lastDataRow})` }, `L${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `L${row}`)
            return
          }
          if (itemType === 'superstructure_builtup_ramps_ramp') {
            spreadsheet.updateCell({ formula: `=C${row}` }, `J${row}`)
            spreadsheet.updateCell({ formula: `=J${row}*H${row}/27` }, `L${row}`)
            spreadsheet.cellFormat({ color: getColumnBColor(row, parsedData) }, `B${row}`)
            return
          }
          if (itemType === 'superstructure_builtup_ramps_ramp_sum') {
            const { firstDataRow, lastDataRow } = formulaInfo
            spreadsheet.updateCell({ formula: `=SUM(J${firstDataRow}:J${lastDataRow})` }, `J${row}`)
            spreadsheet.updateCell({ formula: `=SUM(L${firstDataRow}:L${lastDataRow})` }, `L${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `L${row}`)
            return
          }
          if (itemType === 'superstructure_builtup_stair_knee_wall') {
            spreadsheet.updateCell({ formula: `=C${row}` }, `I${row}`)
            spreadsheet.updateCell({ formula: `=I${row}*H${row}` }, `J${row}`)
            spreadsheet.updateCell({ formula: `=J${row}*G${row}/27` }, `L${row}`)
            spreadsheet.cellFormat({ color: getColumnBColor(row, parsedData) }, `B${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `I${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `L${row}`)
            return
          }
          if (itemType === 'superstructure_builtup_stair_styrofoam') {
            const { takeoffJSumFirstRow, takeoffJSumLastRow, heightRefRow } = formulaInfo
            if (takeoffJSumFirstRow != null && takeoffJSumLastRow != null) {
              spreadsheet.updateCell({ formula: `=SUM(J${takeoffJSumFirstRow}:J${takeoffJSumLastRow})` }, `C${row}`)
            }
            if (heightRefRow != null) spreadsheet.updateCell({ formula: `=H${heightRefRow}` }, `H${row}`)
            spreadsheet.updateCell({ formula: `=C${row}` }, `J${row}`)
            spreadsheet.updateCell({ formula: `=J${row}*H${row}/27` }, `L${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `L${row}`)
            return
          }
          if (itemType === 'superstructure_builtup_stairs') {
            spreadsheet.updateCell({ formula: `=11/12` }, `F${row}`)
            spreadsheet.updateCell({ formula: `=7/12` }, `H${row}`)
            spreadsheet.updateCell({ formula: `=C${row}*G${row}*F${row}` }, `J${row}`)
            spreadsheet.updateCell({ formula: `=J${row}*H${row}/27` }, `L${row}`)
            spreadsheet.updateCell({ formula: `=C${row}` }, `M${row}`)
            spreadsheet.cellFormat({ color: getColumnBColor(row, parsedData) }, `B${row}`)
            return
          }
          if (itemType === 'superstructure_stair_slab') {
            const { takeoffRefRow, widthRefRow, heightValue } = formulaInfo
            if (takeoffRefRow != null) spreadsheet.updateCell({ formula: `=C${takeoffRefRow}*1.3` }, `C${row}`)
            if (widthRefRow != null) spreadsheet.updateCell({ formula: `=G${widthRefRow}` }, `G${row}`)
            if (heightValue != null) spreadsheet.updateCell({ value: heightValue }, `H${row}`)
            spreadsheet.updateCell({ formula: `=C${row}` }, `I${row}`)
            spreadsheet.updateCell({ formula: `=I${row}*H${row}` }, `J${row}`)
            spreadsheet.updateCell({ formula: `=J${row}*G${row}/27` }, `L${row}`)
            spreadsheet.cellFormat({ color: getColumnBColor(row, parsedData) }, `B${row}`)
            return
          }
          if (itemType === 'superstructure_builtup_stair_sum') {
            const { firstDataRow, lastDataRow } = formulaInfo
            spreadsheet.updateCell({ formula: `=SUM(I${firstDataRow}:I${lastDataRow})` }, `I${row}`)
            spreadsheet.updateCell({ formula: `=SUM(J${firstDataRow}:J${lastDataRow})` }, `J${row}`)
            spreadsheet.updateCell({ formula: `=SUM(L${firstDataRow}:L${lastDataRow})` }, `L${row}`)
            spreadsheet.updateCell({ formula: `=SUM(M${firstDataRow}:M${lastDataRow})` }, `M${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `I${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `L${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `M${row}`)
            return
          }
          if (itemType === 'superstructure_concrete_hanger') {
            spreadsheet.updateCell({ formula: `=G${row}*F${row}*C${row}` }, `J${row}`)
            spreadsheet.updateCell({ formula: `=J${row}*H${row}/27` }, `L${row}`)
            spreadsheet.updateCell({ formula: `=C${row}` }, `M${row}`)
            spreadsheet.cellFormat({ color: getColumnBColor(row, parsedData) }, `B${row}`)
            return
          }
          if (itemType === 'superstructure_concrete_hanger_sum') {
            const { firstDataRow, lastDataRow } = formulaInfo
            spreadsheet.updateCell({ formula: `=SUM(J${firstDataRow}:J${lastDataRow})` }, `J${row}`)
            spreadsheet.updateCell({ formula: `=SUM(L${firstDataRow}:L${lastDataRow})` }, `L${row}`)
            spreadsheet.updateCell({ formula: `=SUM(M${firstDataRow}:M${lastDataRow})` }, `M${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `L${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `M${row}`)
            return
          }
          if (itemType === 'superstructure_shear_wall') {
            spreadsheet.updateCell({ formula: `=C${row}` }, `I${row}`)
            spreadsheet.updateCell({ formula: `=I${row}*H${row}` }, `J${row}`)
            spreadsheet.updateCell({ formula: `=J${row}*G${row}/27` }, `L${row}`)
            spreadsheet.cellFormat({ color: getColumnBColor(row, parsedData) }, `B${row}`)
            return
          }
          if (itemType === 'superstructure_shear_walls_sum') {
            const { firstDataRow, lastDataRow } = formulaInfo
            spreadsheet.updateCell({ formula: `=SUM(I${firstDataRow}:I${lastDataRow})` }, `I${row}`)
            spreadsheet.updateCell({ formula: `=SUM(J${firstDataRow}:J${lastDataRow})` }, `J${row}`)
            spreadsheet.updateCell({ formula: `=SUM(L${firstDataRow}:L${lastDataRow})` }, `L${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `I${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `L${row}`)
            return
          }
          if (itemType === 'superstructure_parapet_wall') {
            spreadsheet.updateCell({ formula: `=C${row}` }, `I${row}`)
            spreadsheet.updateCell({ formula: `=I${row}*H${row}` }, `J${row}`)
            spreadsheet.updateCell({ formula: `=J${row}*G${row}/27` }, `L${row}`)
            spreadsheet.cellFormat({ color: getColumnBColor(row, parsedData) }, `B${row}`)
            return
          }
          if (itemType === 'superstructure_parapet_walls_sum') {
            const { firstDataRow, lastDataRow } = formulaInfo
            spreadsheet.updateCell({ formula: `=SUM(I${firstDataRow}:I${lastDataRow})` }, `I${row}`)
            spreadsheet.updateCell({ formula: `=SUM(J${firstDataRow}:J${lastDataRow})` }, `J${row}`)
            spreadsheet.updateCell({ formula: `=SUM(L${firstDataRow}:L${lastDataRow})` }, `L${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `I${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `L${row}`)
            return
          }
          if (itemType === 'superstructure_columns_takeoff') {
            spreadsheet.updateCell({ formula: `=C${row}` }, `M${row}`)
            spreadsheet.cellFormat({ color: getColumnBColor(row, parsedData), textDecoration: 'line-through' }, `B${row}`)
            spreadsheet.cellFormat({ color: '#000000', textDecoration: 'line-through' }, `C${row}`)
            spreadsheet.cellFormat({ color: '#000000', textDecoration: 'line-through' }, `D${row}`)
            spreadsheet.cellFormat({ color: '#000000', textDecoration: 'line-through' }, `M${row}`)
            return
          }
          if (itemType === 'superstructure_columns_final') {
            const { takeoffRefRow } = formulaInfo
            if (takeoffRefRow != null) spreadsheet.updateCell({ formula: `=C${takeoffRefRow}` }, `C${row}`)
            spreadsheet.updateCell({ formula: `=C${row}` }, `M${row}`)
            spreadsheet.cellFormat({ color: getColumnBColor(row, parsedData) }, `B${row}`)
            spreadsheet.cellFormat({ color: '#FF0000', fontWeight: 'normal' }, `I${row}`)
            spreadsheet.cellFormat({ color: '#FF0000', fontWeight: 'normal' }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000', fontWeight: 'normal' }, `L${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `M${row}`)
            return
          }
          if (itemType === 'superstructure_concrete_post') {
            spreadsheet.updateCell({ formula: `=G${row}*H${row}*C${row}` }, `J${row}`)
            spreadsheet.updateCell({ formula: `=J${row}*F${row}/27` }, `L${row}`)
            spreadsheet.updateCell({ formula: `=C${row}` }, `M${row}`)
            spreadsheet.cellFormat({ color: getColumnBColor(row, parsedData) }, `B${row}`)
            return
          }
          if (itemType === 'superstructure_concrete_post_sum') {
            const { firstDataRow, lastDataRow } = formulaInfo
            spreadsheet.updateCell({ formula: `=SUM(J${firstDataRow}:J${lastDataRow})` }, `J${row}`)
            spreadsheet.updateCell({ formula: `=SUM(L${firstDataRow}:L${lastDataRow})` }, `L${row}`)
            spreadsheet.updateCell({ formula: `=SUM(M${firstDataRow}:M${lastDataRow})` }, `M${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `L${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `M${row}`)
            return
          }
          if (itemType === 'superstructure_concrete_encasement') {
            spreadsheet.updateCell({ formula: `=G${row}*H${row}*C${row}` }, `J${row}`)
            spreadsheet.updateCell({ formula: `=J${row}*F${row}/27` }, `L${row}`)
            spreadsheet.updateCell({ formula: `=C${row}` }, `M${row}`)
            spreadsheet.cellFormat({ color: getColumnBColor(row, parsedData) }, `B${row}`)
            return
          }
          if (itemType === 'superstructure_concrete_encasement_sum') {
            const { firstDataRow, lastDataRow } = formulaInfo
            spreadsheet.updateCell({ formula: `=SUM(J${firstDataRow}:J${lastDataRow})` }, `J${row}`)
            spreadsheet.updateCell({ formula: `=SUM(L${firstDataRow}:L${lastDataRow})` }, `L${row}`)
            spreadsheet.updateCell({ formula: `=SUM(M${firstDataRow}:M${lastDataRow})` }, `M${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `L${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `M${row}`)
            return
          }
          if (itemType === 'superstructure_drop_panel_bracket') {
            spreadsheet.updateCell({ formula: `=G${row}*H${row}*C${row}` }, `J${row}`)
            spreadsheet.updateCell({ formula: `=J${row}*F${row}/27` }, `L${row}`)
            spreadsheet.updateCell({ formula: `=C${row}` }, `M${row}`)
            spreadsheet.cellFormat({ color: getColumnBColor(row, parsedData) }, `B${row}`)
            return
          }
          if (itemType === 'superstructure_drop_panel_h') {
            spreadsheet.updateCell({ formula: `=C${row}` }, `J${row}`)
            spreadsheet.updateCell({ formula: `=J${row}*H${row}/27` }, `L${row}`)
            spreadsheet.updateCell({ formula: `=E${row}` }, `M${row}`)
            spreadsheet.cellFormat({ color: getColumnBColor(row, parsedData) }, `B${row}`)
            return
          }
          if (itemType === 'superstructure_drop_panel_sum') {
            const { firstDataRow, lastDataRow } = formulaInfo
            spreadsheet.updateCell({ formula: `=SUM(J${firstDataRow}:J${lastDataRow})` }, `J${row}`)
            spreadsheet.updateCell({ formula: `=SUM(L${firstDataRow}:L${lastDataRow})` }, `L${row}`)
            spreadsheet.updateCell({ formula: `=SUM(M${firstDataRow}:M${lastDataRow})` }, `M${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `L${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `M${row}`)
            return
          }
          if (itemType === 'superstructure_beam') {
            spreadsheet.updateCell({ formula: `=C${row}` }, `I${row}`)
            spreadsheet.updateCell({ formula: `=I${row}*H${row}` }, `J${row}`)
            spreadsheet.updateCell({ formula: `=J${row}*G${row}/27` }, `L${row}`)
            spreadsheet.cellFormat({ color: getColumnBColor(row, parsedData) }, `B${row}`)
            return
          }
          if (itemType === 'superstructure_beams_sum') {
            const { firstDataRow, lastDataRow } = formulaInfo
            spreadsheet.updateCell({ formula: `=SUM(I${firstDataRow}:I${lastDataRow})` }, `I${row}`)
            spreadsheet.updateCell({ formula: `=SUM(J${firstDataRow}:J${lastDataRow})` }, `J${row}`)
            spreadsheet.updateCell({ formula: `=SUM(L${firstDataRow}:L${lastDataRow})` }, `L${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `I${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `L${row}`)
            return
          }
          if (itemType === 'superstructure_curb') {
            spreadsheet.updateCell({ formula: `=C${row}` }, `I${row}`)
            spreadsheet.updateCell({ formula: `=I${row}*H${row}` }, `J${row}`)
            spreadsheet.updateCell({ formula: `=J${row}*G${row}/27` }, `L${row}`)
            spreadsheet.cellFormat({ color: getColumnBColor(row, parsedData) }, `B${row}`)
            return
          }
          if (itemType === 'superstructure_curbs_sum') {
            const { firstDataRow, lastDataRow } = formulaInfo
            spreadsheet.updateCell({ formula: `=SUM(I${firstDataRow}:I${lastDataRow})` }, `I${row}`)
            spreadsheet.updateCell({ formula: `=SUM(J${firstDataRow}:J${lastDataRow})` }, `J${row}`)
            spreadsheet.updateCell({ formula: `=SUM(L${firstDataRow}:L${lastDataRow})` }, `L${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `I${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `L${row}`)
            return
          }
          if (itemType === 'superstructure_concrete_pad') {
            // Check unit from raw data to determine formula application
            const unit = formulaInfo.parsedData?.unit || ''
            const unitLower = String(unit).toLowerCase().trim()
            
            // If unit is SQ FT, apply SQ FT formulas
            if (unitLower.includes('sq') || unitLower === 'sf') {
              spreadsheet.updateCell({ formula: `=C${row}` }, `J${row}`)
              spreadsheet.updateCell({ formula: `=J${row}*H${row}/27` }, `L${row}`)
              // If QTY (column E) exists, M = E, otherwise leave empty
              const hasQty = formulaInfo.parsedData?.parsed?.qty != null
              if (hasQty) {
                spreadsheet.updateCell({ formula: `=E${row}` }, `M${row}`)
              }
              spreadsheet.cellFormat({ color: getColumnBColor(row, parsedData) }, `B${row}`)
            } else {
              // If unit is EA or EACH, apply EA formulas
              spreadsheet.updateCell({ formula: `=C${row}` }, `M${row}`)
              spreadsheet.cellFormat({ color: getColumnBColor(row, parsedData) }, `B${row}`)
            }
            return
          }
          if (itemType === 'superstructure_concrete_pad_sum') {
            const { firstDataRow, lastDataRow } = formulaInfo
            spreadsheet.updateCell({ formula: `=SUM(J${firstDataRow}:J${lastDataRow})` }, `J${row}`)
            spreadsheet.updateCell({ formula: `=SUM(L${firstDataRow}:L${lastDataRow})` }, `L${row}`)
            spreadsheet.updateCell({ formula: `=SUM(M${firstDataRow}:M${lastDataRow})` }, `M${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `L${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `M${row}`)
            return
          }
          if (itemType === 'superstructure_non_shrink_grout') {
            spreadsheet.updateCell({ formula: `=C${row}` }, `M${row}`)
            spreadsheet.cellFormat({ color: getColumnBColor(row, parsedData) }, `B${row}`)
            return
          }
          if (itemType === 'superstructure_non_shrink_grout_sum') {
            const { firstDataRow, lastDataRow } = formulaInfo
            spreadsheet.updateCell({ formula: `=SUM(M${firstDataRow}:M${lastDataRow})` }, `M${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `M${row}`)
            return
          }
          if (itemType === 'superstructure_repair_scope') {
            const subType = formulaInfo.parsedData?.parsed?.itemSubType
            if (subType === 'wall') spreadsheet.updateCell({ formula: `=C${row}` }, `I${row}`)
            else if (subType === 'slab') spreadsheet.updateCell({ formula: `=C${row}` }, `J${row}`)
            else if (subType === 'column') spreadsheet.updateCell({ formula: `=C${row}` }, `M${row}`)
            spreadsheet.cellFormat({ color: '#FF0000', fontWeight: 'normal', fontStyle: 'normal' }, `I${row}`)
            spreadsheet.cellFormat({ color: '#FF0000', fontWeight: 'normal', fontStyle: 'normal' }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000', fontWeight: 'normal', fontStyle: 'normal' }, `M${row}`)
            return
          }
          if (itemType === 'superstructure_repair_scope_sum') {
            const { firstDataRow, lastDataRow } = formulaInfo
            spreadsheet.updateCell({ formula: `=SUM(I${firstDataRow}:I${lastDataRow})` }, `I${row}`)
            spreadsheet.updateCell({ formula: `=SUM(J${firstDataRow}:J${lastDataRow})` }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `I${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `J${row}`)
            return
          }
          if (itemType === 'superstructure_extra_sqft') {
            spreadsheet.updateCell({ formula: `=C${row}` }, `J${row}`)
            spreadsheet.updateCell({ formula: `=J${row}*H${row}/27` }, `L${row}`)
            spreadsheet.cellFormat({ color: getColumnBColor(row, parsedData) }, `B${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `L${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `M${row}`)
            return
          }
          if (itemType === 'superstructure_extra_ft') {
            spreadsheet.updateCell({ formula: `=C${row}*G${row}` }, `J${row}`)
            spreadsheet.updateCell({ formula: `=J${row}*H${row}/27` }, `L${row}`)
            spreadsheet.cellFormat({ color: getColumnBColor(row, parsedData) }, `B${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `L${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `M${row}`)
            return
          }
          if (itemType === 'superstructure_extra_ea') {
            spreadsheet.updateCell({ formula: `=C${row}*F${row}*G${row}` }, `J${row}`)
            spreadsheet.updateCell({ formula: `=J${row}*H${row}/27` }, `L${row}`)
            spreadsheet.updateCell({ formula: `=C${row}` }, `M${row}`)
            spreadsheet.cellFormat({ color: getColumnBColor(row, parsedData) }, `B${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `L${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `M${row}`)
            return
          }
        } else if (section === 'bpp_alternate') {
          // B.P.P. Alternate #2 scope formulas
          if (itemType === 'bpp_street_header') {
            // Street header row - black text, underlined, bold
            spreadsheet.cellFormat({ fontWeight: 'bold', textDecoration: 'underline', color: '#000000' }, `B${row}`)
            return
          }
          if (itemType === 'bpp_gravel') {
            // Gravel: col B black text, col H empty, col I empty, col J = C, CY = J * H / 27
            spreadsheet.cellFormat({ color: '#000000' }, `B${row}`)
            spreadsheet.updateCell({ formula: `=C${row}` }, `J${row}`)
            spreadsheet.updateCell({ formula: `=J${row}*H${row}/27` }, `L${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `L${row}`)
            return
          }
          if (itemType === 'bpp_concrete_sidewalk' || itemType === 'bpp_concrete_driveway') {
            // SQ FT items: col I empty, col J = C, CY = J * H / 27
            const parsed = parsedData || formulaInfo
            const heightFormula = parsed?.parsed?.heightFormula
            if (heightFormula) {
              spreadsheet.updateCell({ formula: `=${heightFormula}` }, `H${row}`)
            }
            spreadsheet.updateCell({ formula: `=C${row}` }, `J${row}`)
            spreadsheet.updateCell({ formula: `=J${row}*H${row}/27` }, `L${row}`)
            return
          }
          if (itemType === 'bpp_concrete_curb' || itemType === 'bpp_concrete_flush_curb') {
            // FT items: FT = C, SQ_FT = I * H, CY = J * G / 27
            const parsed = parsedData || formulaInfo
            const widthFormula = parsed?.parsed?.widthFormula
            if (widthFormula) {
              spreadsheet.updateCell({ formula: `=${widthFormula}` }, `G${row}`)
            }
            spreadsheet.updateCell({ formula: `=C${row}` }, `I${row}`)
            spreadsheet.updateCell({ formula: `=I${row}*H${row}` }, `J${row}`)
            spreadsheet.updateCell({ formula: `=J${row}*G${row}/27` }, `L${row}`)
            return
          }
          if (itemType === 'bpp_expansion_joint') {
            // Expansion joint: FT = C
            spreadsheet.updateCell({ formula: `=C${row}` }, `I${row}`)
            return
          }
          if (itemType === 'bpp_conc_road_base') {
            // Conc road base: black text, col H empty, SQ_FT = C, CY = J * H / 27
            spreadsheet.updateCell({ formula: `=C${row}` }, `J${row}`)
            spreadsheet.updateCell({ formula: `=J${row}*H${row}/27` }, `L${row}`)
            spreadsheet.cellFormat({ color: '#000000' }, `B${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `L${row}`)
            return
          }
          if (itemType === 'bpp_full_depth_asphalt') {
            // Full depth asphalt: col I empty, col J = C, CY = J * H / 27
            const parsed = parsedData || formulaInfo
            const heightFormula = parsed?.parsed?.heightFormula
            if (heightFormula) {
              spreadsheet.updateCell({ formula: `=${heightFormula}` }, `H${row}`)
            }
            spreadsheet.updateCell({ formula: `=C${row}` }, `J${row}`)
            spreadsheet.updateCell({ formula: `=J${row}*H${row}/27` }, `L${row}`)
            return
          }
          if (itemType === 'bpp_sum') {
            const { firstDataRow, lastDataRow, sumColumns } = formulaInfo
            // Sum should be in J column, not I (for sidewalk, driveway, asphalt items)
            if (sumColumns && sumColumns.includes('I')) {
              spreadsheet.updateCell({ formula: `=SUM(I${firstDataRow}:I${lastDataRow})` }, `I${row}`)
              spreadsheet.cellFormat({ color: '#FF0000' }, `I${row}`)
            }
            if (sumColumns && sumColumns.includes('J')) {
              spreadsheet.updateCell({ formula: `=SUM(J${firstDataRow}:J${lastDataRow})` }, `J${row}`)
              spreadsheet.cellFormat({ color: '#FF0000' }, `J${row}`)
            }
            if (sumColumns && sumColumns.includes('L')) {
              spreadsheet.updateCell({ formula: `=SUM(L${firstDataRow}:L${lastDataRow})` }, `L${row}`)
              spreadsheet.cellFormat({ color: '#FF0000' }, `L${row}`)
            }
            return
          }
        } else if (section === 'civil_sitework') {
          // Civil / Sitework formulas
          if (itemType === 'civil_demo_asphalt') {
            // Demo asphalt: J = C, L = J * H / 27
            spreadsheet.updateCell({ formula: `=C${row}` }, `J${row}`)
            spreadsheet.updateCell({ formula: `=J${row}*H${row}/27` }, `L${row}`)
            return
          }
          if (itemType === 'civil_demo_curb') {
            // Demo curb: I = C, J = I * H, L = J * G / 27
            spreadsheet.updateCell({ formula: `=C${row}` }, `I${row}`)
            spreadsheet.updateCell({ formula: `=I${row}*H${row}` }, `J${row}`)
            spreadsheet.updateCell({ formula: `=J${row}*G${row}/27` }, `L${row}`)
            return
          }
          if (itemType === 'civil_demo_fence') {
            // Demo fence: I = C, J = I * H
            spreadsheet.updateCell({ formula: `=C${row}` }, `I${row}`)
            spreadsheet.updateCell({ formula: `=I${row}*H${row}` }, `J${row}`)
            return
          }
          if (itemType === 'civil_demo_wall') {
            // Demo wall: I = C, J = I * H, L = J * G / 27
            spreadsheet.updateCell({ formula: `=C${row}` }, `I${row}`)
            spreadsheet.updateCell({ formula: `=I${row}*H${row}` }, `J${row}`)
            spreadsheet.updateCell({ formula: `=J${row}*G${row}/27` }, `L${row}`)
            return
          }
          if (itemType === 'civil_demo_pipe' || itemType === 'civil_demo_rail') {
            // Demo pipe/rail: I = C
            spreadsheet.updateCell({ formula: `=C${row}` }, `I${row}`)
            return
          }
          if (itemType === 'civil_demo_ea') {
            // Demo EA items (sign, manhole, fire hydrant, utility pole, valve, inlet): M = C
            spreadsheet.updateCell({ formula: `=C${row}` }, `M${row}`)
            return
          }
          if (itemType === 'civil_demo_sum') {
            const { firstDataRow, lastDataRow, sumColumns } = formulaInfo
            if (sumColumns && sumColumns.includes('I')) {
              spreadsheet.updateCell({ formula: `=SUM(I${firstDataRow}:I${lastDataRow})` }, `I${row}`)
              spreadsheet.cellFormat({ color: '#FF0000' }, `I${row}`)
            }
            if (sumColumns && sumColumns.includes('J')) {
              spreadsheet.updateCell({ formula: `=SUM(J${firstDataRow}:J${lastDataRow})` }, `J${row}`)
              spreadsheet.cellFormat({ color: '#FF0000' }, `J${row}`)
            }
            if (sumColumns && sumColumns.includes('L')) {
              spreadsheet.updateCell({ formula: `=SUM(L${firstDataRow}:L${lastDataRow})` }, `L${row}`)
              spreadsheet.cellFormat({ color: '#FF0000' }, `L${row}`)
            }
            if (sumColumns && sumColumns.includes('M')) {
              spreadsheet.updateCell({ formula: `=SUM(M${firstDataRow}:M${lastDataRow})` }, `M${row}`)
              spreadsheet.cellFormat({ color: '#FF0000' }, `M${row}`)
            }
            return
          }
          // Excavation items
          // Height is manual input (empty), Col I empty, Col J = C for SQ FT or G*F*C for EA
          if (itemType === 'civil_exc_transformer' || itemType === 'civil_exc_sidewalk') {
            // SQ FT items: J = C, L = J*H/27
            spreadsheet.updateCell({ formula: `=C${row}` }, `J${row}`)
            spreadsheet.updateCell({ formula: `=J${row}*H${row}/27` }, `L${row}`)
            return
          }
          if (itemType === 'civil_exc_bollard') {
            // EA items: F = SQRT(3.14*1.5*1.5/4), G = SQRT(3.14*1.5*1.5/4), H = empty, J = G*F*C, L = J*H/27
            spreadsheet.updateCell({ formula: `=SQRT(3.14*1.5*1.5/4)` }, `F${row}`)
            spreadsheet.updateCell({ formula: `=SQRT(3.14*1.5*1.5/4)` }, `G${row}`)
            spreadsheet.updateCell({ formula: `=G${row}*F${row}*C${row}` }, `J${row}`)
            spreadsheet.updateCell({ formula: `=J${row}*H${row}/27` }, `L${row}`)
            return
          }
          if (itemType === 'civil_exc_sum') {
            const { firstDataRow, lastDataRow } = formulaInfo
            spreadsheet.updateCell({ formula: `=SUM(J${firstDataRow}:J${lastDataRow})` }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `J${row}`)
            spreadsheet.updateCell({ formula: `=SUM(L${firstDataRow}:L${lastDataRow})*1.25` }, `L${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `L${row}`)
            return
          }
          // Gravel items
          if (itemType === 'civil_gravel_item') {
            // I empty, J = C (red), L = J*H/27 (red with yellow background)
            spreadsheet.updateCell({ formula: `=C${row}` }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `J${row}`)
            spreadsheet.updateCell({ formula: `=J${row}*H${row}/27` }, `L${row}`)
            spreadsheet.cellFormat({ color: '#FF0000', backgroundColor: '#FFF2CC' }, `L${row}`)
            return
          }
          if (itemType === 'civil_gravel_sum') {
            const { firstDataRow, lastDataRow } = formulaInfo
            spreadsheet.updateCell({ formula: `=SUM(J${firstDataRow}:J${lastDataRow})` }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `J${row}`)
            spreadsheet.updateCell({ formula: `=SUM(L${firstDataRow}:L${lastDataRow})` }, `L${row}`)
            spreadsheet.cellFormat({ color: '#FF0000', backgroundColor: '#FFF2CC' }, `L${row}`)
            return
          }
          // Concrete Pavement items
          if (itemType === 'civil_concrete_pavement') {
            // I empty, J = C, L = J*H/27
            spreadsheet.updateCell({ formula: `=C${row}` }, `J${row}`)
            spreadsheet.updateCell({ formula: `=J${row}*H${row}/27` }, `L${row}`)
            return
          }
          if (itemType === 'civil_concrete_pavement_sum') {
            const { firstDataRow, lastDataRow } = formulaInfo
            spreadsheet.updateCell({ formula: `=SUM(J${firstDataRow}:J${lastDataRow})` }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `J${row}`)
            spreadsheet.updateCell({ formula: `=SUM(L${firstDataRow}:L${lastDataRow})` }, `L${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `L${row}`)
            return
          }
          // Asphalt items
          if (itemType === 'civil_asphalt') {
            // Height formula, I empty, J = C, L = J*H/27
            const parsed = parsedData || formulaInfo
            const heightFormula = parsed?.parsed?.heightFormula
            if (heightFormula) {
              spreadsheet.updateCell({ formula: `=${heightFormula}` }, `H${row}`)
            }
            spreadsheet.updateCell({ formula: `=C${row}` }, `J${row}`)
            spreadsheet.updateCell({ formula: `=J${row}*H${row}/27` }, `L${row}`)
            return
          }
          if (itemType === 'civil_asphalt_sum') {
            const { firstDataRow, lastDataRow } = formulaInfo
            spreadsheet.updateCell({ formula: `=SUM(J${firstDataRow}:J${lastDataRow})` }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `J${row}`)
            spreadsheet.updateCell({ formula: `=SUM(L${firstDataRow}:L${lastDataRow})` }, `L${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `L${row}`)
            return
          }
          // Pads items
          if (itemType === 'civil_pads') {
            // I empty, J = C, L = J*H/27, M = E (qty)
            spreadsheet.updateCell({ formula: `=C${row}` }, `J${row}`)
            spreadsheet.updateCell({ formula: `=J${row}*H${row}/27` }, `L${row}`)
            spreadsheet.updateCell({ formula: `=E${row}` }, `M${row}`)
            return
          }
          if (itemType === 'civil_pads_sum') {
            const { firstDataRow, lastDataRow } = formulaInfo
            spreadsheet.updateCell({ formula: `=SUM(J${firstDataRow}:J${lastDataRow})` }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `J${row}`)
            spreadsheet.updateCell({ formula: `=SUM(L${firstDataRow}:L${lastDataRow})` }, `L${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `L${row}`)
            spreadsheet.updateCell({ formula: `=SUM(M${firstDataRow}:M${lastDataRow})` }, `M${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `M${row}`)
            return
          }
          // Soil Erosion items - all values in col I, J, L, M should be red
          if (itemType === 'civil_soil_stabilized') {
            // I empty, J = C (red), L = J*H/27 (red)
            spreadsheet.updateCell({ formula: `=C${row}` }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `J${row}`)
            spreadsheet.updateCell({ formula: `=J${row}*H${row}/27` }, `L${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `L${row}`)
            return
          }
          if (itemType === 'civil_soil_silt_fence') {
            // I = C (red), J = I*H (red)
            spreadsheet.updateCell({ formula: `=C${row}` }, `I${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `I${row}`)
            spreadsheet.updateCell({ formula: `=I${row}*H${row}` }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `J${row}`)
            return
          }
          if (itemType === 'civil_soil_inlet_filter') {
            // M = C (red)
            spreadsheet.updateCell({ formula: `=C${row}` }, `M${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `M${row}`)
            return
          }
          // Fence items
          if (itemType === 'civil_fence') {
            // I = C, J = I*H
            spreadsheet.updateCell({ formula: `=C${row}` }, `I${row}`)
            spreadsheet.updateCell({ formula: `=I${row}*H${row}` }, `J${row}`)
            return
          }
          if (itemType === 'civil_fence_sum') {
            const { firstDataRow, lastDataRow } = formulaInfo
            spreadsheet.updateCell({ formula: `=SUM(I${firstDataRow}:I${lastDataRow})` }, `I${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `I${row}`)
            spreadsheet.updateCell({ formula: `=SUM(J${firstDataRow}:J${lastDataRow})` }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `J${row}`)
            return
          }
          // Bollard items
          // Bollard items
          if (itemType === 'civil_bollard_footing_item') {
            const parsed = parsedData || formulaInfo
            const dims = parsed?.parsed?.bollardDimensions
            // Use footing height if available, otherwise 4
            const hVal = dims?.footingH || 4

            // F, G = SQRT(3.14*1.5*1.5/4) -> 1.5 is 18" footing diameter
            const sideFormula = `SQRT(3.14*1.5*1.5/4)`
            spreadsheet.updateCell({ formula: `=${sideFormula}` }, `F${row}`)
            spreadsheet.updateCell({ formula: `=${sideFormula}` }, `G${row}`)

            // H = derived from name (mentioned in footing)
            spreadsheet.updateCell({ value: hVal }, `H${row}`)

            // J = G*F*C
            spreadsheet.updateCell({ formula: `=G${row}*F${row}*C${row}` }, `J${row}`)

            // L = J*H/27
            spreadsheet.updateCell({ formula: `=J${row}*H${row}/27` }, `L${row}`)

            // M = C
            spreadsheet.updateCell({ formula: `=C${row}` }, `M${row}`)
            return
          }

          if (itemType === 'civil_bollard_footing_sum') {
            const { firstDataRow, lastDataRow } = formulaInfo
            // Sum J, L, M
            spreadsheet.updateCell({ formula: `=SUM(J${firstDataRow}:J${lastDataRow})` }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `J${row}`)

            spreadsheet.updateCell({ formula: `=SUM(L${firstDataRow}:L${lastDataRow})` }, `L${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `L${row}`)

            spreadsheet.updateCell({ formula: `=SUM(M${firstDataRow}:M${lastDataRow})` }, `M${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `M${row}`)
            return
          }

          if (itemType === 'civil_bollard_simple_item') {
            // M = C
            spreadsheet.updateCell({ formula: `=C${row}` }, `M${row}`)
            return
          }

          if (itemType === 'civil_bollard_simple_sum') {
            const { firstDataRow, lastDataRow } = formulaInfo
            // Sum M
            spreadsheet.updateCell({ formula: `=SUM(M${firstDataRow}:M${lastDataRow})` }, `M${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `M${row}`)
            return
          }

          // Site items (Hydrant, Wheel stop, Drain, Protection, Signages, Main line)
          if (itemType === 'civil_site_item') {
            // M = C
            spreadsheet.updateCell({ formula: `=C${row}` }, `M${row}`)
            return
          }

          if (itemType === 'civil_site_sum') {
            const { firstDataRow, lastDataRow } = formulaInfo
            // Sum M
            spreadsheet.updateCell({ formula: `=SUM(M${firstDataRow}:M${lastDataRow})` }, `M${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `M${row}`)
            return
          }

          // Drains & Utilities and Alternate items
          if (itemType === 'civil_drains_utilities_item' || itemType === 'civil_alternate_item') {
            const unit = parsedData?.unit || 'EA'
            // Particulars in conditional color based on takeoff
            spreadsheet.cellFormat({ color: getColumnBColor(row, parsedData) }, `B${row}`)

            if (unit === 'FT' || unit === 'LF') {
              // I = C
              spreadsheet.updateCell({ formula: `=C${row}` }, `I${row}`)
              spreadsheet.cellFormat({ color: '#FF0000' }, `I${row}`)
            } else {
              // M = C
              spreadsheet.updateCell({ formula: `=C${row}` }, `M${row}`)
              spreadsheet.cellFormat({ color: '#FF0000' }, `M${row}`)
            }
            return
          }

          // Alternate Header
          if (itemType === 'civil_alternate_header') {
            spreadsheet.cellFormat({ backgroundColor: '#E2EFDA' }, `B${row}`)
            return
          }

          // Ele subsection items (Excavation, Backfill, Gravel)
          if (itemType === 'civil_ele_item') {
            const { subSubsectionName, takeoffSourceType } = formulaInfo

            // Find the source row for takeoff
            let sourceRow = null
            let hasFormulaInC = false

            if (takeoffSourceType === 'drains_conduit') {
              // Find "Proposed underground electrical conduit" in Drains & Utilities
              for (let i = 0; i < calculationData.length; i++) {
                const rowData = calculationData[i]
                const particulars = rowData[1] ? String(rowData[1]).toLowerCase() : ''
                if (particulars.includes('proposed underground electrical conduit')) {
                  sourceRow = i + 1
                  break
                }
              }
            } else if (takeoffSourceType === 'ele_excavation') {
              // Find Excavation item in Ele subsection (the data row, not the header)
              for (let i = 0; i < calculationData.length; i++) {
                const rowData = calculationData[i]
                const particulars = rowData[1] ? String(rowData[1]).toLowerCase() : ''
                // Look for the Excavation data row (not the header which has "  Excavation:")
                if (particulars === 'excavation' && i > 0) {
                  // Make sure this is in the Ele section by checking previous rows
                  let inEleSection = false
                  for (let j = i - 1; j >= 0; j--) {
                    const prevRow = calculationData[j]
                    if (prevRow[1] && String(prevRow[1]).trim() === 'Ele:') {
                      inEleSection = true
                      break
                    }
                    // Stop if we hit another subsection header
                    if (prevRow[1] && String(prevRow[1]).endsWith(':') && !String(prevRow[1]).startsWith('  ')) {
                      break
                    }
                  }
                  if (inEleSection) {
                    sourceRow = i + 1
                    break
                  }
                }
              }
            }

            // Apply formulas
            if (sourceRow) {
              // C = reference to source row's C column (black text)
              spreadsheet.updateCell({ formula: `=C${sourceRow}` }, `C${row}`)
              spreadsheet.cellFormat({ color: '#000000' }, `C${row}`)
              hasFormulaInC = true
            }

            // Item name in column B - use black since C has formula or no data
            spreadsheet.cellFormat({ color: hasFormulaInC ? '#000000' : getColumnBColor(row, parsedData) }, `B${row}`)

            // I = empty (no formula)

            // Check if this is Gravel item
            const isGravel = subSubsectionName === 'Gravel'

            if (isGravel) {
              // Gravel: J = C*G (black text)
              spreadsheet.updateCell({ formula: `=C${row}*G${row}` }, `J${row}`)
              spreadsheet.cellFormat({ color: '#000000' }, `J${row}`)

              // Gravel: L = J*H/27 (black text, yellow background)
              spreadsheet.updateCell({ formula: `=J${row}*H${row}/27` }, `L${row}`)
              spreadsheet.cellFormat({ color: '#000000', backgroundColor: '#FFF2CC' }, `L${row}`)
            } else {
              // Excavation & Backfill: J = H*C (black text)
              spreadsheet.updateCell({ formula: `=H${row}*C${row}` }, `J${row}`)
              spreadsheet.cellFormat({ color: '#000000' }, `J${row}`)

              // Excavation & Backfill: L = J*G/27 (black text)
              spreadsheet.updateCell({ formula: `=J${row}*G${row}/27` }, `L${row}`)
              spreadsheet.cellFormat({ color: '#000000' }, `L${row}`)
            }

            return
          }

          if (itemType === 'civil_ele_sum') {
            const { firstDataRow, lastDataRow, subSubsectionName } = formulaInfo

            // Check if this is Gravel sum or Excavation sum
            const isGravel = subSubsectionName === 'Gravel'
            const isExcavation = subSubsectionName === 'Excavation'

            // Sum J (SQ FT) - red text
            spreadsheet.updateCell({ formula: `=SUM(J${firstDataRow}:J${lastDataRow})` }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `J${row}`)

            // Sum L (CY) - red text, yellow background for Gravel, multiply by 1.25 for Excavation
            if (isExcavation) {
              spreadsheet.updateCell({ formula: `=SUM(L${firstDataRow}:L${lastDataRow})*1.25` }, `L${row}`)
            } else {
              spreadsheet.updateCell({ formula: `=SUM(L${firstDataRow}:L${lastDataRow})` }, `L${row}`)
            }
            if (isGravel) {
              spreadsheet.cellFormat({ color: '#FF0000', backgroundColor: '#FFF2CC' }, `L${row}`)
            } else {
              spreadsheet.cellFormat({ color: '#FF0000' }, `L${row}`)
            }

            return
          }

          // Gas and Water subsection items (Excavation, Backfill, Gravel)
          if (itemType === 'civil_gas_water_item') {
            const { subsectionName, subSubsectionName, takeoffSourceType, isGravel } = formulaInfo

            // Find the source row for takeoff
            let sourceRow = null
            let hasFormulaInC = false

            if (takeoffSourceType === 'drains_gas_lateral') {
              // Find "Proposed Underground gas service lateral" in Drains & Utilities
              // For Gas subsection, Drains & Utilities is BELOW, so search forward from current row
              const searchStart = subsectionName === 'Gas' ? row : 0
              for (let i = searchStart; i < calculationData.length; i++) {
                const rowData = calculationData[i]
                const particulars = rowData[1] ? String(rowData[1]).toLowerCase() : ''
                if (particulars.includes('proposed') && particulars.includes('gas service lateral')) {
                  // Make sure this is in Drains & Utilities section
                  let inDrainsUtilities = false
                  for (let j = i - 1; j >= 0; j--) {
                    const prevRow = calculationData[j]
                    if (prevRow[1] && String(prevRow[1]).includes('Drains & Utilities:')) {
                      inDrainsUtilities = true
                      break
                    }
                    if (prevRow[1] && String(prevRow[1]).endsWith(':') && !String(prevRow[1]).startsWith('  ')) {
                      break
                    }
                  }
                  if (inDrainsUtilities) {
                    sourceRow = i + 1
                    break
                  }
                }
              }
            } else if (takeoffSourceType === 'drains_water_main') {
              // Find "Proposed Underground water main" in Drains & Utilities
              // For Gas subsection, Drains & Utilities is BELOW, so search forward from current row
              const searchStart = subsectionName === 'Gas' ? row : 0
              for (let i = searchStart; i < calculationData.length; i++) {
                const rowData = calculationData[i]
                const particulars = rowData[1] ? String(rowData[1]).toLowerCase() : ''
                if (particulars.includes('proposed') && particulars.includes('water main')) {
                  // Make sure this is in Drains & Utilities section
                  let inDrainsUtilities = false
                  for (let j = i - 1; j >= 0; j--) {
                    const prevRow = calculationData[j]
                    if (prevRow[1] && String(prevRow[1]).includes('Drains & Utilities:')) {
                      inDrainsUtilities = true
                      break
                    }
                    if (prevRow[1] && String(prevRow[1]).endsWith(':') && !String(prevRow[1]).startsWith('  ')) {
                      break
                    }
                  }
                  if (inDrainsUtilities) {
                    sourceRow = i + 1
                    break
                  }
                }
              }
            } else if (takeoffSourceType === 'drains_sanitary_sewer') {
              // Find "Proposed Underground 8" PVC sanitary sewer service" in Drains & Utilities
              for (let i = 0; i < calculationData.length; i++) {
                const rowData = calculationData[i]
                const particulars = rowData[1] ? String(rowData[1]).toLowerCase() : ''
                if (particulars.includes('proposed') && particulars.includes('sanitary sewer')) {
                  sourceRow = i + 1
                  break
                }
              }
            } else if (takeoffSourceType === 'gas_excavation') {
              // Gas Backfill references Gas Excavation item
              // The structure is always:
              // Row N: Excavation header
              // Row N+1: Excavation data (source)
              // Row N+2: Excavation sum
              // Row N+3: Blank
              // Row N+4: Backfill header
              // Row N+5: Backfill data (current row)
              // So the Excavation data is always 4 rows above the Backfill data
              sourceRow = row - 4
            } else if (takeoffSourceType === 'water_backfill_5_rows_above') {
              // Water Backfill references Water Excavation item 5 rows above
              // The structure is:
              // Row N: Excavation header
              // Row N+1: Excavation item 1 (source)
              // Row N+2: Excavation item 2
              // Row N+3: Excavation sum
              // Row N+4: Blank
              // Row N+5: Backfill header
              // Row N+6: Backfill data (current row)
              // So the Excavation item 1 is always 5 rows above the Backfill data
              sourceRow = row - 5
            } else if (takeoffSourceType === 'water_excavation_item1') {
              // Find first Excavation item in Water subsection
              for (let i = 0; i < calculationData.length; i++) {
                const rowData = calculationData[i]
                const particulars = rowData[1] ? String(rowData[1]).toLowerCase() : ''
                if (particulars.includes('proposed') && particulars.includes('water service lateral') && i > 0) {
                  // Make sure this is in the Water Excavation section
                  let inWaterExcavation = false
                  for (let j = i - 1; j >= 0; j--) {
                    const prevRow = calculationData[j]
                    if (prevRow[1] && String(prevRow[1]).trim() === '  Excavation:') {
                      // Check if this is under Water subsection
                      for (let k = j - 1; k >= 0; k--) {
                        const subsecRow = calculationData[k]
                        if (subsecRow[1] && String(subsecRow[1]).trim() === 'Water:') {
                          inWaterExcavation = true
                          break
                        }
                        if (subsecRow[1] && String(subsecRow[1]).endsWith(':') && !String(subsecRow[1]).startsWith('  ')) {
                          break
                        }
                      }
                      break
                    }
                  }
                  if (inWaterExcavation) {
                    sourceRow = i + 1
                    break
                  }
                }
              }
            }

            // Apply formulas
            if (sourceRow) {
              // C = reference to source row's C column (black text)
              spreadsheet.updateCell({ formula: `=C${sourceRow}` }, `C${row}`)
              spreadsheet.cellFormat({ color: '#000000' }, `C${row}`)
              hasFormulaInC = true
            }

            // Item name in column B - use black since C has formula or no data
            spreadsheet.cellFormat({ color: hasFormulaInC ? '#000000' : getColumnBColor(row, parsedData) }, `B${row}`)

            // I = empty (no formula)

            if (isGravel) {
              // Gravel items: J = C*G (black text)
              spreadsheet.updateCell({ formula: `=C${row}*G${row}` }, `J${row}`)
              spreadsheet.cellFormat({ color: '#000000' }, `J${row}`)

              // Gravel items: L = J*H/27 (black text, yellow background)
              spreadsheet.updateCell({ formula: `=J${row}*H${row}/27` }, `L${row}`)
              spreadsheet.cellFormat({ color: '#000000', backgroundColor: '#FFF2CC' }, `L${row}`)
            } else {
              // Excavation & Backfill: J = H*C (black text)
              spreadsheet.updateCell({ formula: `=H${row}*C${row}` }, `J${row}`)
              spreadsheet.cellFormat({ color: '#000000' }, `J${row}`)

              // Excavation & Backfill: L = J*G/27 (black text)
              spreadsheet.updateCell({ formula: `=J${row}*G${row}/27` }, `L${row}`)
              spreadsheet.cellFormat({ color: '#000000' }, `L${row}`)
            }

            return
          }

          if (itemType === 'civil_gas_water_sum') {
            const { firstDataRow, lastDataRow, isGravel, subSubsectionName } = formulaInfo

            // Check if this is Excavation sum
            const isExcavation = subSubsectionName === 'Excavation'

            // Sum J (SQ FT) - red text
            spreadsheet.updateCell({ formula: `=SUM(J${firstDataRow}:J${lastDataRow})` }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `J${row}`)

            // Sum L (CY) - red text, yellow background for Gravel, multiply by 1.25 for Excavation
            if (isExcavation) {
              spreadsheet.updateCell({ formula: `=SUM(L${firstDataRow}:L${lastDataRow})*1.25` }, `L${row}`)
            } else {
              spreadsheet.updateCell({ formula: `=SUM(L${firstDataRow}:L${lastDataRow})` }, `L${row}`)
            }
            if (isGravel) {
              spreadsheet.cellFormat({ color: '#FF0000', backgroundColor: '#FFF2CC' }, `L${row}`)
            } else {
              spreadsheet.cellFormat({ color: '#FF0000' }, `L${row}`)
            }

            return
          }
        }

        // Manual Superstructure Items (CIP Stairs and Stairs - Infilled tads)

        if (section === 'superstructure') {
          // CIP Stairs - Header (underlined, no other data)
          if (itemType === 'superstructure_manual_cip_stairs_header') {
            spreadsheet.cellFormat({ textDecoration: 'underline' }, `B${row}`)
            return
          }

          // CIP Stairs - Landings: I=empty, J=C, K=empty, L=J*H/27
          if (itemType === 'superstructure_manual_cip_stairs_landing') {
            // J = C
            spreadsheet.updateCell({ formula: `=C${row}` }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `J${row}`)
            // L = J*H/27
            spreadsheet.updateCell({ formula: `=J${row}*H${row}/27` }, `L${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `L${row}`)
            return
          }

          // CIP Stairs - Stairs: J=C*G*F, K=empty, L=J*H/27, M=C
          if (itemType === 'superstructure_manual_cip_stairs_stair') {
            // J = C*G*F
            spreadsheet.updateCell({ formula: `=C${row}*G${row}*F${row}` }, `J${row}`)
            // L = J*H/27
            spreadsheet.updateCell({ formula: `=J${row}*H${row}/27` }, `L${row}`)
            // M = C
            spreadsheet.updateCell({ formula: `=C${row}` }, `M${row}`)
            return
          }

          // CIP Stairs - Stair Slab: I=C, J=I*H/27, K=empty, L=J*G/27, M=empty
          if (itemType === 'superstructure_manual_cip_stairs_slab') {
            const { stairsRowNum, slabCMultiplier } = formulaInfo
            // C = reference to stairs row C (with optional multiplier)
            if (slabCMultiplier && slabCMultiplier !== 1) {
              spreadsheet.updateCell({ formula: `=C${stairsRowNum}*${slabCMultiplier}` }, `C${row}`)
            } else {
              spreadsheet.updateCell({ formula: `=C${stairsRowNum}` }, `C${row}`)
            }
            // I = C
            spreadsheet.updateCell({ formula: `=C${row}` }, `I${row}`)
            // J = I*H
            spreadsheet.updateCell({ formula: `=I${row}*H${row}` }, `J${row}`)
            // L = J*G/27
            spreadsheet.updateCell({ formula: `=J${row}*G${row}/27` }, `L${row}`)
            return
          }

          // CIP Stairs - Sum: SUM of I, J, L, M (excludes Landings row)
          if (itemType === 'superstructure_manual_cip_stairs_sum') {
            const { firstDataRow, lastDataRow } = formulaInfo
            // Sum I
            spreadsheet.updateCell({ formula: `=SUM(I${firstDataRow}:I${lastDataRow})` }, `I${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `I${row}`)
            // Sum J
            spreadsheet.updateCell({ formula: `=SUM(J${firstDataRow}:J${lastDataRow})` }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `J${row}`)
            // Sum L
            spreadsheet.updateCell({ formula: `=SUM(L${firstDataRow}:L${lastDataRow})` }, `L${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `L${row}`)
            // Sum M
            spreadsheet.updateCell({ formula: `=SUM(M${firstDataRow}:M${lastDataRow})` }, `M${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `M${row}`)
            return
          }

          // Stairs - Infilled tads - Header (underlined, no other data)
          if (itemType === 'superstructure_manual_infilled_header') {
            spreadsheet.cellFormat({ textDecoration: 'underline' }, `B${row}`)
            return
          }

          // Stairs - Infilled tads - Landing 1: J=C, L=(J*H)/27
          if (itemType === 'superstructure_manual_infilled_landing_1') {
            // J = C
            spreadsheet.updateCell({ formula: `=C${row}` }, `J${row}`)
            // L = (J*H)/27
            spreadsheet.updateCell({ formula: `=J${row}*H${row}/27` }, `L${row}`)
            // H = blue background
            spreadsheet.cellFormat({ backgroundColor: '#D9E1F2' }, `H${row}`)
            return
          }

          // Stairs - Infilled tads - Landing 2: C=C{landing1}, H=1.5/12 (blue bg), J=C, L=(J*H)/27/2
          if (itemType === 'superstructure_manual_infilled_landing_2') {
            const { landing1RowNum } = formulaInfo
            // C = reference to Landing 1 C
            spreadsheet.updateCell({ formula: `=C${landing1RowNum}` }, `C${row}`)
            // H = blue background (no formula)
            spreadsheet.updateCell({ formula: `` }, `H${row}`)
            spreadsheet.cellFormat({ backgroundColor: '#D9E1F2' }, `H${row}`)
            // J = C
            spreadsheet.updateCell({ formula: `=C${row}` }, `J${row}`)
            // L = (J*H)/27/2
            spreadsheet.updateCell({ formula: `=J${row}*H${row}/27/2` }, `L${row}`)
            return
          }

          // Stairs - Infilled tads - Landing Sum: J=J{landing1}, L=SUM (all RED)
          if (itemType === 'superstructure_manual_infilled_landing_sum') {
            const { landing1RowNum, firstDataRow, lastDataRow } = formulaInfo
            // J = J{landing1Row}
            spreadsheet.updateCell({ formula: `=J${landing1RowNum}` }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `J${row}`)
            // L = SUM(L range)
            spreadsheet.updateCell({ formula: `=SUM(L${firstDataRow}:L${lastDataRow})` }, `L${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `L${row}`)
            return
          }

          // Stairs - Infilled tads - Stairs: J=C*G*F, L=J*H/27, M=C (All RED)
          if (itemType === 'superstructure_manual_infilled_stair') {
            // J = C*G*F
            spreadsheet.updateCell({ formula: `=C${row}*G${row}*F${row}` }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `J${row}`)
            // L = J*H/27
            spreadsheet.updateCell({ formula: `=J${row}*H${row}/27` }, `L${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `L${row}`)
            // M = C
            spreadsheet.updateCell({ formula: `=C${row}` }, `M${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `M${row}`)
            return
          }
        }

        // Apply generic formulas if we got them from generators
        if (formulas) {
          if (formulas.qty) spreadsheet.updateCell({ formula: `=${formulas.qty}` }, `E${row}`)
          if (formulas.ft) spreadsheet.updateCell({ formula: `=${formulas.ft}` }, `I${row}`)
          if (formulas.sqFt) spreadsheet.updateCell({ formula: `=${formulas.sqFt}` }, `J${row}`)
          if (formulas.lbs) {
            spreadsheet.updateCell({ formula: `=${formulas.lbs}` }, `K${row}`)
            if (section === 'excavation' && parsedData?.subsection !== 'backfill') {
              spreadsheet.cellFormat({ textDecoration: 'line-through' }, `K${row}`)
            }
          }
          if (formulas.cy) spreadsheet.updateCell({ formula: `=${formulas.cy}` }, `L${row}`)
          if (formulas.qtyFinal) spreadsheet.updateCell({ formula: `=${formulas.qtyFinal}` }, `M${row}`)
        }
      } catch (error) {
        console.error(`Error applying formula at row ${row}:`, error)
      }
    })

    // Apply deferred foundation sum L formulas
    deferredFoundationSumL.forEach(({ row: r, lFormula }) => {
      try {
        spreadsheet.updateCell({ formula: lFormula }, `L${r}`)
      } catch (e) {
        console.error(`Error applying deferred Foundation sum L at row ${r}:`, e)
      }
    })

    // Apply formatting
    try {
      // Format header row - Estimate column (A) has yellow background
      spreadsheet.cellFormat({ fontWeight: 'bold', textAlign: 'center', backgroundColor: '#FFFF00', color: '#000000' }, 'A1')
      spreadsheet.cellFormat({ fontWeight: 'bold', textAlign: 'center', color: '#000000' }, 'B1:M1')

      // Format section and subsection headers
      calculationData.forEach((row, rowIndex) => {
        const rowNum = rowIndex + 1

        if (row[0] && !row[1]) {
          const sectionName = String(row[0])
          let backgroundColor = '#F4B084'
          if (sectionName === 'Demolition') backgroundColor = '#E5B7AF'
          else if (['Excavation', 'Rock Excavation', 'SOE', 'Foundation', 'Waterproofing', 'Trenching', 'Superstructure', 'B.P.P. Alternate #2 scope', 'Civil / Sitework'].includes(sectionName)) {
            backgroundColor = '#C6E0B4'
          }
          spreadsheet.cellFormat({ fontWeight: 'bold', backgroundColor, fontSize: '11pt' }, `A${rowNum}:M${rowNum}`)
          if (sectionName === 'Foundation' || sectionName === 'Trenching') {
            spreadsheet.cellFormat({ fontWeight: 'normal', backgroundColor, fontSize: '11pt' }, `C${rowNum}:D${rowNum}`)
          }
        }

        if (!row[0] && row[1]) {
          const bContent = String(row[1])
          if (bContent.endsWith(':') || bContent.startsWith('  ')) {
            if (bContent.includes('For demo Extra line item use this') || bContent.includes('For Backfill Extra line item use this') || bContent.includes('For soil excavation Extra line item use this') || bContent.includes('For rock excavation Extra line item use this') || bContent.includes('For foundation Extra line item use this') || bContent.includes('For Superstructure Extra line item use this')) {
              spreadsheet.cellFormat({ fontWeight: 'bold', fontStyle: 'italic', backgroundColor: '#FFFF00' }, `B${rowNum}`)
            } else if (bContent.includes('Line drill:') || bContent.includes('Underpinning:')) {
              spreadsheet.cellFormat({ fontWeight: 'bold', fontStyle: 'italic', backgroundColor: '#E2EFDA' }, `B${rowNum}`)
            } else {
              spreadsheet.cellFormat({ fontWeight: 'bold', fontStyle: 'italic' }, `B${rowNum}`)
            }
          } else if (bContent !== 'Havg' && bContent !== 'Gravel' && bContent !== 'Conc road base' && !bContent.startsWith('Street name:')) {
            spreadsheet.cellFormat({ color: '#FF0000' }, `B${rowNum}`)
          }
        }
      })

      // Column formatting
      spreadsheet.cellFormat({ textAlign: 'center' }, 'A:A')
      const numColumns = ['E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M']
      numColumns.forEach(col => {
        try { spreadsheet.numberFormat('0.00', `${col}:${col}`) } catch (e) { }
      })
    } catch (error) {
      console.error('Error applying formatting:', error)
    }
  }

  // Constants for auto-save timing
  const DEBOUNCE_DELAY = 2000 // Wait 2 seconds after last change before saving
  const MAX_WAIT_TIME = 10000 // Force save after 10 seconds of continuous changes
  const MAX_RETRY_COUNT = 3
  const RETRY_DELAY = 2000

  // Extract all images from the spreadsheet for separate storage
  // Syncfusion's saveAsJson() doesn't include images, so we need to extract them manually
  const extractImages = useCallback(() => {
    if (!spreadsheetRef.current) return []

    const spreadsheet = spreadsheetRef.current
    const images = []

    try {
      // Access sheets from the spreadsheet model
      const sheets = spreadsheet.sheets

      if (!sheets || !Array.isArray(sheets)) {
        return images
      }

      sheets.forEach((sheet, sheetIndex) => {
        try {
          // Try both 'image' and 'images' property names
          const sheetImages = sheet.image || sheet.images || []

          if (sheetImages && Array.isArray(sheetImages) && sheetImages.length > 0) {
            sheetImages.forEach((image, idx) => {
              if (image && image.src) {
                images.push({
                  sheetIndex,
                  imageId: image.id || `image_${sheetIndex}_${idx}_${Date.now()}`,
                  src: image.src,
                  top: image.top || 0,
                  left: image.left || 0,
                  width: image.width,
                  height: image.height,
                })
              }
            })
          }
        } catch (sheetError) {
          // Silently handle errors for individual sheets
        }
      })
    } catch (error) {
      console.error('Error extracting images:', error)
    }

    return images
  }, [])

  // Restore images to the spreadsheet after loading from JSON
  const restoreImages = useCallback((savedImages) => {
    if (!spreadsheetRef.current || !savedImages || savedImages.length === 0) return

    const spreadsheet = spreadsheetRef.current

    // Small delay to ensure spreadsheet is fully loaded
    setTimeout(() => {
      savedImages.forEach((image) => {
        try {
          // Build the image model for Syncfusion
          const imageModel = [{
            src: image.src,
            id: image.imageId,
            top: image.top || 0,
            left: image.left || 0,
            width: image.width,
            height: image.height,
          }]

          // Insert the image into the correct sheet
          spreadsheet.insertImage(imageModel, `A1`, image.sheetIndex)
        } catch (error) {
          // Silently handle image restoration errors
        }
      })
    }, 500)
  }, [])

  // Core save function with retry logic
  const saveSpreadsheet = useCallback(async (isManual = false) => {
    if (!spreadsheetRef.current || !proposalId) return false

    // Prevent concurrent saves
    if (isSavingRef.current) {
      pendingSaveRef.current = true
      return false
    }

    isSavingRef.current = true
    setIsSaving(true)
    setSaveError(null)

    try {
      // Get the spreadsheet JSON (includes all sheets: Calculations Sheet + Proposal Sheet)
      const json = await spreadsheetRef.current.saveAsJson()
      const jsonObject = json.jsonObject || json

      // Dev: verify both sheets are in the save payload
      if (import.meta.env.DEV) {
        const sheets = jsonObject?.Workbook?.sheets || []
        const sheetNames = sheets.map(s => s.name || s.Name).filter(Boolean)
        if (!sheetNames.includes('Proposal Sheet')) {
          console.warn('[ProposalDetail] Save payload missing Proposal Sheet. Sheets:', sheetNames)
        }
      }

      // Extract images separately since saveAsJson doesn't include them
      const images = extractImages()

      // Save both spreadsheet JSON and images
      await proposalAPI.update(proposalId, {
        spreadsheetJson: jsonObject,
        images: images,
        unusedRawDataRows: unusedRawDataRowsRef.current // Sync current unused rows state
      })

      const now = new Date()
      setLastSaved(now)
      lastSaveTimeRef.current = now
      setHasUnsavedChanges(false)
      retryCountRef.current = 0

      return true
    } catch (error) {
      console.error('Error saving spreadsheet:', error)
      setSaveError(error.message || 'Failed to save')

      // Retry logic for failed saves
      if (retryCountRef.current < MAX_RETRY_COUNT) {
        retryCountRef.current++
        setTimeout(() => {
          saveSpreadsheet(false)
        }, RETRY_DELAY)
      } else {
        // Show error only after all retries failed
        toast.error('Failed to save changes. Please check your connection.', {
          duration: 5000,
        })
      }
      return false
    } finally {
      isSavingRef.current = false
      setIsSaving(false)

      // If there was a pending save request, process it
      if (pendingSaveRef.current) {
        pendingSaveRef.current = false
        setTimeout(() => {
          markDirtyAndScheduleSave()
        }, 500)
      }
    }
  }, [proposalId, extractImages])

  // Mark as dirty and schedule a save
  const markDirtyAndScheduleSave = useCallback(() => {
    setHasUnsavedChanges(true)
    setSaveError(null)

    // Clear existing debounce timeout
    if (saveTimeoutRef.current) {
      clearTimeout(saveTimeoutRef.current)
    }

    // Set debounce timeout - save after user stops making changes
    saveTimeoutRef.current = setTimeout(() => {
      saveSpreadsheet(false)
      // Clear max wait timeout since we're saving now
      if (maxWaitTimeoutRef.current) {
        clearTimeout(maxWaitTimeoutRef.current)
        maxWaitTimeoutRef.current = null
      }
    }, DEBOUNCE_DELAY)

    // Set max wait timeout - force save if user keeps making continuous changes
    if (!maxWaitTimeoutRef.current) {
      maxWaitTimeoutRef.current = setTimeout(() => {
        // Clear debounce timeout and save immediately
        if (saveTimeoutRef.current) {
          clearTimeout(saveTimeoutRef.current)
          saveTimeoutRef.current = null
        }
        saveSpreadsheet(false)
        maxWaitTimeoutRef.current = null
      }, MAX_WAIT_TIME)
    }
  }, [saveSpreadsheet])

  // Cleanup timeouts on unmount and save any pending changes
  useEffect(() => {
    return () => {
      if (saveTimeoutRef.current) clearTimeout(saveTimeoutRef.current)
      if (maxWaitTimeoutRef.current) clearTimeout(maxWaitTimeoutRef.current)
    }
  }, [])

  // Warn user before leaving with unsaved changes
  useEffect(() => {
    const handleBeforeUnload = (e) => {
      if (hasUnsavedChanges) {
        e.preventDefault()
        e.returnValue = 'You have unsaved changes. Are you sure you want to leave?'
        // Try to save before leaving
        if (spreadsheetRef.current && proposalId) {
          saveSpreadsheet(false)
        }
        return e.returnValue
      }
    }

    window.addEventListener('beforeunload', handleBeforeUnload)
    return () => window.removeEventListener('beforeunload', handleBeforeUnload)
  }, [hasUnsavedChanges, proposalId, saveSpreadsheet])

  // Handle Ctrl+S to trigger immediate save
  useEffect(() => {
    const handleKeyDown = (e) => {
      if ((e.ctrlKey || e.metaKey) && e.key === 's') {
        e.preventDefault()

        // Clear pending timeouts and save immediately
        if (saveTimeoutRef.current) clearTimeout(saveTimeoutRef.current)
        if (maxWaitTimeoutRef.current) clearTimeout(maxWaitTimeoutRef.current)
        maxWaitTimeoutRef.current = null

        if (hasUnsavedChanges) {
          saveSpreadsheet(true)
        }

        toast('✌🏻 You are good to go: your changes are being autosaved', {
          duration: 3000,
          position: 'bottom-right',
          style: {
            fontSize: '13px',
          },
          icon: null,
        })
      }
    }

    window.addEventListener('keydown', handleKeyDown)
    return () => window.removeEventListener('keydown', handleKeyDown)
  }, [hasUnsavedChanges, saveSpreadsheet])

  // Handle spreadsheet cell save event
  const handleCellSave = useCallback(() => {
    markDirtyAndScheduleSave()
  }, [markDirtyAndScheduleSave])

  // Handle spreadsheet action complete event - captures most changes
  const handleActionComplete = useCallback((args) => {
    // Actions that should NOT trigger a save (read-only operations)
    const noSaveActions = [
      'copy', 'select', 'scroll', 'zoom',
      'find', 'getData', 'refresh'
    ]
    // Note: We DO save on 'gotoSheet' - when user switches away from Proposal Sheet,
    // we need to capture any pending edits (cell may have been saved but we want to persist)
    // Save on any action that's not in the no-save list
    if (args.action && !noSaveActions.includes(args.action)) {
      markDirtyAndScheduleSave()
    }
  }, [markDirtyAndScheduleSave])

  // Handle before cell save - captures formula and value changes
  const handleBeforeCellSave = useCallback(() => {
    // This fires before the cell is saved
  }, [])

  // Handle cell edit event - captures when editing starts
  const handleCellEdit = useCallback(() => {
    // We'll save when editing completes via cellSave
  }, [])

  // Handle before cell update for format changes
  const handleBeforeCellUpdate = useCallback(() => {
    markDirtyAndScheduleSave()
  }, [markDirtyAndScheduleSave])

  // Fallback: Periodically check for image changes
  // This catches image insertions that might not trigger actionComplete
  const lastImageCountRef = useRef(0)

  useEffect(() => {
    if (!spreadsheetRef.current || !hasLoadedFromJson.current) return

    const checkForImageChanges = () => {
      try {
        const images = extractImages()
        const currentCount = images.length

        if (currentCount !== lastImageCountRef.current) {
          lastImageCountRef.current = currentCount
          markDirtyAndScheduleSave()
        }
      } catch (error) {
        // Ignore errors during image check
      }
    }

    // Check for image changes every 2 seconds
    const intervalId = setInterval(checkForImageChanges, 2000)

    return () => clearInterval(intervalId)
  }, [extractImages, markDirtyAndScheduleSave])

  const handleSettingsSave = async (settingsData) => {
    try {
      await proposalAPI.update(proposalId, settingsData)
      setProposal(prev => ({ ...prev, ...settingsData }))
      toast.success('Settings updated successfully')
    } catch (error) {
      console.error('Error updating settings:', error)
      toast.error('Error updating settings')
      throw error
    }
  }

  const handleUpdateUnusedRowStatus = async (rowIndex, isUsed) => {
    try {
      // Optimistic update
      const updatedRows = proposal.unusedRawDataRows.map(row => {
        if (row.rowIndex === rowIndex) {
          return { ...row, isUsed }
        }
        return row
      })
      setProposal(prev => ({
        ...prev,
        unusedRawDataRows: updatedRows
      }))

      // API call
      // The original code used proposalAPI.updateUnusedRowStatus, which is a wrapper.
      // Assuming the new implementation uses axios directly as shown in the instruction.
      // If proposalAPI.updateUnusedRowStatus is still desired, this part needs adjustment.
      // For now, I'll use the axios call as provided in the instruction.
      // Note: `id` and `unusedRows` are not defined in this scope.
      // I'll assume `proposalId` for `id` and `proposal.unusedRawDataRows` for `unusedRows`
      // and that `axios` is imported.
      const token = localStorage.getItem('token')
      const config = {
        headers: {
          Authorization: `Bearer ${token}`
        }
      }

      await proposalAPI.updateUnusedRowStatus(proposalId, rowIndex, isUsed)

      // No need to fetch proposal again as we optimistcally updated
    } catch (error) {
      console.error('Error updating row status:', error)
      toast.error('Failed to update row status')
      // Revert on error by fetching
      // fetchProposal()
    }
  }

  const handleBulkUpdateUnusedRowStatus = async (updates) => {
    try {
      // Optimistic update
      const updatedRows = [...proposal.unusedRawDataRows]
      updates.forEach(({ rowIndex, isUsed }) => {
        const row = updatedRows.find(r => r.rowIndex === rowIndex)
        if (row) {
          row.isUsed = isUsed
        }
      })
      setProposal(prev => ({
        ...prev,
        unusedRawDataRows: updatedRows
      }))

      // API call
      await proposalAPI.updateUnusedRowStatusBulk(proposalId, updates)

      toast.success('Changes saved successfully')
    } catch (error) {
      console.error('Error bulk updating row status:', error)
      toast.error('Failed to save changes')
      // Revert on error by fetching
      // fetchProposal()
    }
  }

  // Handle navigation away - save first if there are unsaved changes
  const handleNavigateBack = useCallback(async () => {
    if (hasUnsavedChanges) {
      // Clear pending timeouts
      if (saveTimeoutRef.current) clearTimeout(saveTimeoutRef.current)
      if (maxWaitTimeoutRef.current) clearTimeout(maxWaitTimeoutRef.current)

      // Show saving indicator
      toast.loading('Saving changes...', { id: 'nav-save' })

      try {
        await saveSpreadsheet(true)
        toast.success('Changes saved!', { id: 'nav-save', duration: 1500 })
      } catch (error) {
        toast.error('Failed to save, navigating anyway...', { id: 'nav-save', duration: 1500 })
      }
    }

    navigate('/proposals')
  }, [hasUnsavedChanges, saveSpreadsheet, navigate])

  if (isLoading) {
    return (
      <div className="flex h-screen bg-gray-50">
        <Sidebar collapsed={sidebarCollapsed} onToggle={toggleSidebar} />
        <div className="flex-1 flex flex-col overflow-hidden">
          <TopBar />
          <div className="flex-1 flex items-center justify-center">
            <p className="text-gray-500">Loading proposal...</p>
          </div>
        </div>
      </div>
    )
  }

  if (!proposal) {
    return (
      <div className="flex h-screen bg-gray-50">
        <Sidebar collapsed={sidebarCollapsed} onToggle={toggleSidebar} />
        <div className="flex-1 flex flex-col overflow-hidden">
          <TopBar />
          <div className="flex-1 flex items-center justify-center">
            <p className="text-gray-500">Proposal not found</p>
          </div>
        </div>
      </div>
    )
  }

  return (
    <div className="flex h-screen bg-gray-50">
      <Sidebar collapsed={sidebarCollapsed} onToggle={toggleSidebar} />

      <div className="flex-1 flex flex-col overflow-hidden">
        <TopBar
          hasUnsavedChanges={hasUnsavedChanges}
          leftContent={
            <div className="flex items-center gap-4">
              <button
                onClick={handleNavigateBack}
                className="p-2 -ml-2 text-gray-500 hover:text-gray-700 hover:bg-gray-100 rounded-full transition-colors"
                title="Back to Proposals"
              >
                <FiArrowLeft size={20} />
              </button>
              <div>
                <h1 className="text-lg font-semibold text-gray-900">{proposal.name}</h1>
                <p className="text-sm text-gray-500 flex items-center gap-2">
                  <span>{proposal.client} • {proposal.project}</span>
                  <span className="text-gray-300">|</span>
                  {isSaving ? (
                    <span className="flex items-center gap-1.5 text-blue-600">
                      <span className="inline-block w-2 h-2 bg-blue-600 rounded-full animate-pulse"></span>
                      Saving...
                    </span>
                  ) : saveError ? (
                    <span className="flex items-center gap-1.5 text-red-600">
                      <span className="inline-block w-2 h-2 bg-red-600 rounded-full"></span>
                      Save failed - retrying...
                    </span>
                  ) : hasUnsavedChanges ? (
                    <span className="flex items-center gap-1.5 text-amber-600">
                      <span className="inline-block w-2 h-2 bg-amber-500 rounded-full"></span>
                      Unsaved changes
                    </span>
                  ) : lastSaved ? (
                    <span className="flex items-center gap-1.5 text-green-600">
                      <span className="inline-block w-2 h-2 bg-green-500 rounded-full"></span>
                      Saved {lastSaved.toLocaleTimeString()}
                    </span>
                  ) : (
                    <span className="text-gray-400">Ready</span>
                  )}
                </p>
              </div>
            </div>
          }
          actionButtons={
            <>
              <button
                onClick={() => setIsUnusedDataModalOpen(true)}
                className="flex items-center gap-2 px-4 py-2 text-sm font-medium text-gray-700 bg-white border border-gray-300 rounded-lg hover:bg-gray-50 transition-colors"
              >
                <FiCheckSquare size={18} />
                Unused Data
              </button>
              <button
                onClick={() => setIsPreviewModalOpen(true)}
                className="flex items-center gap-2 px-4 py-2 text-sm font-medium text-gray-700 bg-white border border-gray-300 rounded-lg hover:bg-gray-50 transition-colors"
              >
                <FiEye size={18} />
                Preview Raw Data
              </button>
              <button
                onClick={() => setIsSettingsModalOpen(true)}
                className="flex items-center gap-2 px-4 py-2 text-sm font-medium text-gray-700 bg-white border border-gray-300 rounded-lg hover:bg-gray-50 transition-colors"
              >
                <FiSettings size={18} />
                Settings
              </button>
            </>
          }
        />

        {/* Spreadsheet */}
        <div className="flex-1 overflow-hidden relative">
          {isSpreadsheetLoading && (
            <div className="absolute inset-0 bg-white bg-opacity-90 flex items-center justify-center z-50">
              <div className="text-center">
                <div className="inline-block animate-spin rounded-full h-10 w-10 border-4 border-blue-500 border-t-transparent mb-4"></div>
                <p className="text-gray-700 font-medium">Generating spreadsheet...</p>
                <p className="text-gray-500 text-sm mt-1">This may take a moment for large files</p>
              </div>
            </div>
          )}
          <SpreadsheetComponent
            ref={spreadsheetRef}
            allowSave={true}
            saveUrl="https://krmestimators.com/dotnet/api/spreadsheet/save"
            showFormulaBar={true}
            showRibbon={true}
            showSheetTabs={true}
            allowEditing={true}
            allowUndoRedo={true}
            allowImage={true}
            enableClipboard={true}
            // Cell events
            cellSave={handleCellSave}
            cellEdit={handleCellEdit}
            beforeCellSave={handleBeforeCellSave}
            beforeCellUpdate={handleBeforeCellUpdate}
            // Action events - captures all modifications including images
            actionComplete={handleActionComplete}
            // File menu events - trigger save after file menu operations
            // fileMenuItemSelect={() => setTimeout(() => markDirtyAndScheduleSave(), 1000)}
            // Created event
            created={() => {
              // Fallback: when spreadsheet is ready and we have generated data but proposal wasn't built yet (timing)
              const calcData = calculationDataRef.current
              if (calcData.length > 0 && !proposalBuiltRef.current && spreadsheetRef.current && !proposal?.spreadsheetJson) {
                try {
                  buildProposalSheet(spreadsheetRef.current, {
                    calculationData: calcData,
                    formulaData: formulaDataRef.current,
                    rockExcavationTotals: rockExcavationTotalsRef.current,
                    lineDrillTotalFT: lineDrillTotalFTRef.current,
                    rawData: rawDataRef.current
                  })
                  proposalBuiltRef.current = true
                } catch (e) {
                  console.error('Error building proposal sheet (onCreated):', e)
                }
              }
            }}
          >
            <SheetsDirective>
              <SheetDirective name="Proposal Sheet" />
              <SheetDirective name="Calculations Sheet">
                <ColumnsDirective>
                  {generateColumnConfigs().map((config, idx) => (
                    <ColumnDirective key={idx} width={config.width} />
                  ))}
                </ColumnsDirective>
              </SheetDirective>
            </SheetsDirective>
          </SpreadsheetComponent>
        </div>
      </div>

      {/* Modals */}
      <RawDataPreviewModal
        isOpen={isPreviewModalOpen}
        onClose={() => setIsPreviewModalOpen(false)}
        rawExcelData={proposal.rawExcelData}
        proposalId={proposal?._id}
        onSaveSuccess={async () => {
          if (!proposal?._id) return
          try {
            const response = await proposalAPI.getById(proposal._id)
            setProposal(response.proposal)
            needReapplyAfterRawSave.current = true
            setIsPreviewModalOpen(false)
            toast.success('Raw data saved. Rebuilding sheets…')
          } catch (e) {
            toast.error('Failed to refresh after save')
          }
        }}
        onRebuildWithRawData={(editedRawExcelData) => {
          setProposal(prev => prev ? { ...prev, rawExcelData: editedRawExcelData } : prev)
          needReapplyAfterRawSave.current = true
          setIsPreviewModalOpen(false)
          toast.success('Rebuilding calculation and proposal sheets from edited raw data…')
        }}
      />

      <UnusedRawDataModal
        isOpen={isUnusedDataModalOpen}
        onClose={() => setIsUnusedDataModalOpen(false)}
        unusedRows={proposal.unusedRawDataRows}
        onUpdateRowStatus={handleUpdateUnusedRowStatus}
        onBulkUpdateRowStatus={handleBulkUpdateUnusedRowStatus}
        headers={proposal.rawExcelData?.headers}
      />

      <ProposalSettingsModal
        isOpen={isSettingsModalOpen}
        onClose={() => setIsSettingsModalOpen(false)}
        proposal={proposal}
        onSave={handleSettingsSave}
      />
    </div >
  )
}

export default ProposalDetail