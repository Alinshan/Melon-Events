import React, { useState, useEffect, useMemo } from 'react'
import {
  LayoutDashboard,
  Receipt,
  Package,
  Settings,
  Plus,
  Download,
  Printer,
  Trash2,
  CheckCircle,
  Wine,
  Calendar,
  IndianRupee,
  Briefcase,
  Wallet,
  Droplets,
  AlertTriangle,
  Edit2,
  Sun,
  Moon,
  Info,
  ExternalLink,
  RotateCcw
} from 'lucide-react'
import { motion, AnimatePresence } from 'framer-motion'
import * as XLSX from 'xlsx'
import { jsPDF } from 'jspdf'
import autoTable from 'jspdf-autotable'
import 'jspdf-autotable'

// --- Types ---
interface ItemSettings {
  platePrice: number
  glassPrice: number
  plateCleaningPrice: number
  glassCleaningPrice: number
  totalPlates: number
  totalGlasses: number
  employees?: string[]
  owner1Name?: string
  owner2Name?: string
  owner1Investment?: number
  owner2Investment?: number
  plateLossPrice?: number
  glassLossPrice?: number
}

interface Rental {
  id: string
  date: string
  customerName: string
  plateCount: number
  glassCount: number
  platePrice: number // captured at time of rental
  glassPrice: number
  discount: number
  travelExpense: number
  total: number
  isReturned: boolean
  returnDate?: string
  damageCharge?: number
  isCleaned?: boolean
  cleaningCost?: number
  unusedPlateCount?: number
  unusedGlassCount?: number
  platesCleaned?: number
  glassesCleaned?: number
}

interface SalaryRecord {
  id: string
  date: string
  employeeName: string
  rentalId: string
  plateCount: number
  glassCount: number
  plateRate: number
  glassRate: number
  totalSalary: number
}

interface ExpenseRecord {
  id: string
  date: string
  category: string
  description: string
  amount: number
}

interface OwnerTransaction {
  id: string
  date: string
  from: string
  to: string
  amount: number
  notes: string
  subType?: 'investment' | 'profit' | 'transfer' // For BUSINESS -> OWNER payouts
}

interface AppData {
  settings: ItemSettings
  rentals: Rental[]
  salaries: SalaryRecord[]
  expenses: ExpenseRecord[]
  ownerTransactions: OwnerTransaction[]
}

const EditRentalModal = ({ rental, settings, onSave, onCancel }: { rental: Rental, settings: any, onSave: (r: Rental) => void, onCancel: () => void }) => {
  const [edited, setEdited] = useState(rental)

  const handleSubmit = (e: React.FormEvent) => {
    e.preventDefault()
    const usedPlates = Math.max(0, edited.plateCount - (edited.unusedPlateCount || 0))
    const usedGlasses = Math.max(0, edited.glassCount - (edited.unusedGlassCount || 0))
    const total = (usedPlates * (edited.platePrice || settings.platePrice)) + (usedGlasses * (edited.glassPrice || settings.glassPrice)) - (edited.discount || 0)
    onSave({ ...edited, total })
  }

  return (
    <div style={{
      position: 'fixed', top: 0, left: 0, right: 0, bottom: 0,
      background: 'rgba(15, 23, 42, 0.8)', backdropFilter: 'blur(4px)',
      display: 'flex', alignItems: 'center', justifyContent: 'center', zIndex: 1000
    }}>
      <form className="card" style={{ width: '600px', border: '1px solid var(--border-color)' }} onSubmit={handleSubmit}>
        <h3>Edit Rental {rental.id}</h3>
        <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '1rem', marginBottom: '1.5rem' }}>
          <div className="form-group">
            <label>Customer Name</label>
            <input type="text" value={edited.customerName} onChange={e => setEdited({ ...edited, customerName: e.target.value })} required />
          </div>
          <div className="form-group">
            <label>Date</label>
            <input type="date" value={edited.date} onChange={e => setEdited({ ...edited, date: e.target.value })} required />
          </div>
          <div className="form-group">
            <label>Plate Count</label>
            <input type="number" value={edited.plateCount} onChange={e => setEdited({ ...edited, plateCount: Number(e.target.value) })} required />
          </div>
          <div className="form-group">
            <label>Glass Count</label>
            <input type="number" value={edited.glassCount} onChange={e => setEdited({ ...edited, glassCount: Number(e.target.value) })} required />
          </div>
          <div className="form-group">
            <label>Unused Plates</label>
            <input type="number" value={edited.unusedPlateCount || 0} onChange={e => setEdited({ ...edited, unusedPlateCount: Number(e.target.value) })} />
          </div>
          <div className="form-group">
            <label>Unused Glasses</label>
            <input type="number" value={edited.unusedGlassCount || 0} onChange={e => setEdited({ ...edited, unusedGlassCount: Number(e.target.value) })} />
          </div>
          <div className="form-group">
            <label>Travel Expense (₹)</label>
            <input type="number" value={edited.travelExpense || 0} onChange={e => setEdited({ ...edited, travelExpense: Number(e.target.value) })} />
          </div>
          <div className="form-group">
            <label>Discount (₹)</label>
            <input type="number" value={edited.discount || 0} onChange={e => setEdited({ ...edited, discount: Number(e.target.value) })} />
          </div>
          <div className="form-group">
            <label>Plate Price (₹)</label>
            <input type="number" step="0.01" value={edited.platePrice} onChange={e => setEdited({ ...edited, platePrice: Number(e.target.value) })} />
          </div>
          <div className="form-group">
            <label>Glass Price (₹)</label>
            <input type="number" step="0.01" value={edited.glassPrice} onChange={e => setEdited({ ...edited, glassPrice: Number(e.target.value) })} />
          </div>
          <div className="form-group">
            <label>Damage Charge (₹)</label>
            <input type="number" value={edited.damageCharge || 0} onChange={e => setEdited({ ...edited, damageCharge: Number(e.target.value) })} />
          </div>
          {edited.isReturned && (
            <div className="form-group">
              <label>Return Date</label>
              <input type="date" value={edited.returnDate || ''} onChange={e => setEdited({ ...edited, returnDate: e.target.value })} required />
            </div>
          )}
        </div>
        <div style={{ display: 'flex', gap: '1rem', justifyContent: 'flex-end' }}>
          <button type="button" className="btn-secondary" onClick={onCancel}>Cancel</button>
          <button type="submit" className="btn-primary">Save Changes</button>
        </div>
      </form>
    </div>
  )
}

const EditSalaryModal = ({ salary, onSave, onCancel }: { salary: SalaryRecord, onSave: (s: SalaryRecord) => void, onCancel: () => void }) => {
  const [edited, setEdited] = useState(salary)

  const handleSubmit = (e: React.FormEvent) => {
    e.preventDefault()
    const totalSalary = (edited.plateCount * edited.plateRate) + (edited.glassCount * edited.glassRate)
    onSave({ ...edited, totalSalary })
  }

  return (
    <div style={{
      position: 'fixed', top: 0, left: 0, right: 0, bottom: 0,
      background: 'rgba(15, 23, 42, 0.8)', backdropFilter: 'blur(4px)',
      display: 'flex', alignItems: 'center', justifyContent: 'center', zIndex: 1000
    }}>
      <form className="card" style={{ width: '400px', border: '1px solid var(--border-color)' }} onSubmit={handleSubmit}>
        <h3>Edit Salary Record</h3>
        <div className="form-group">
          <label>Employee Name</label>
          <input type="text" value={edited.employeeName} onChange={e => setEdited({ ...edited, employeeName: e.target.value })} required />
        </div>
        <div className="form-group">
          <label>Date</label>
          <input type="date" value={edited.date} onChange={e => setEdited({ ...edited, date: e.target.value })} required />
        </div>
        <div className="form-group">
          <label>Plate Count</label>
          <input type="number" value={edited.plateCount} onChange={e => setEdited({ ...edited, plateCount: Number(e.target.value) })} required />
        </div>
        <div className="form-group">
          <label>Glass Count</label>
          <input type="number" value={edited.glassCount} onChange={e => setEdited({ ...edited, glassCount: Number(e.target.value) })} required />
        </div>
        <div style={{ display: 'flex', gap: '1rem', marginTop: '1.5rem', justifyContent: 'flex-end' }}>
          <button type="button" className="btn-secondary" onClick={onCancel}>Cancel</button>
          <button type="submit" className="btn-primary">Save Changes</button>
        </div>
      </form>
    </div>
  )
}

const EditExpenseModal = ({ expense, onSave, onCancel }: { expense: ExpenseRecord, onSave: (e: ExpenseRecord) => void, onCancel: () => void }) => {
  const [edited, setEdited] = useState(expense)

  return (
    <div style={{
      position: 'fixed', top: 0, left: 0, right: 0, bottom: 0,
      background: 'rgba(15, 23, 42, 0.8)', backdropFilter: 'blur(4px)',
      display: 'flex', alignItems: 'center', justifyContent: 'center', zIndex: 1000
    }}>
      <form className="card" style={{ width: '400px', border: '1px solid var(--border-color)' }} onSubmit={(e) => { e.preventDefault(); onSave(edited); }}>
        <h3>Edit Expense</h3>
        <div className="form-group">
          <label>Date</label>
          <input type="date" value={edited.date} onChange={e => setEdited({ ...edited, date: e.target.value })} required />
        </div>
        <div className="form-group">
          <label>Category</label>
          <input type="text" value={edited.category} onChange={e => setEdited({ ...edited, category: e.target.value })} required />
        </div>
        <div className="form-group">
          <label>Description</label>
          <input type="text" value={edited.description} onChange={e => setEdited({ ...edited, description: e.target.value })} required />
        </div>
        <div className="form-group">
          <label>Amount (₹)</label>
          <input type="number" value={edited.amount} onChange={e => setEdited({ ...edited, amount: Number(e.target.value) })} required />
        </div>
        <div style={{ display: 'flex', gap: '1rem', marginTop: '1.5rem', justifyContent: 'flex-end' }}>
          <button type="button" className="btn-secondary" onClick={onCancel}>Cancel</button>
          <button type="submit" className="btn-primary">Save Changes</button>
        </div>
      </form>
    </div>
  )
}

// --- App Component ---
function App() {
  const [theme, setTheme] = useState<'dark' | 'light'>('dark')
  const [activeTab, setActiveTab] = useState<'dashboard' | 'rentals' | 'inventory' | 'salaries' | 'expenses' | 'owner-profit' | 'about'>('dashboard')
  const [data, setData] = useState<AppData>({
    settings: { 
      platePrice: 4, 
      glassPrice: 1, 
      plateCleaningPrice: 1, 
      glassCleaningPrice: 0.4, 
      totalPlates: 1000, 
      totalGlasses: 500, 
      employees: [],
      owner1Name: 'ASHRAF M',
      owner2Name: 'RASHEED M',
      owner1Investment: 0,
      owner2Investment: 0
    },
    rentals: [],
    salaries: [],
    expenses: [],
    ownerTransactions: []
  })
  const [loading, setLoading] = useState(true)
  const [cleaningEmployee, setCleaningEmployee] = useState('')
  const [newEmployeeName, setNewEmployeeName] = useState('')
  const [manualPlates, setManualPlates] = useState<string | number>('')
  const [manualGlasses, setManualGlasses] = useState<string | number>('')
  const [manualCleaningDate, setManualCleaningDate] = useState(new Date().toISOString().split('T')[0])
  const [manualEmployee, setManualEmployee] = useState('')
  const [returningRentalId, setReturningRentalId] = useState<string | null>(null)
  const [damageChargeInput, setDamageChargeInput] = useState<string>('')
  const [unusedPlatesInput, setUnusedPlatesInput] = useState<string>('')
  const [unusedGlassesInput, setUnusedGlassesInput] = useState<string>('')
  const [cleaningRentalId, setCleaningRentalId] = useState<string | null>(null)
  const [cleaningEmployeeSelect, setCleaningEmployeeSelect] = useState<string>('')
  const [cleaningPlatesInput, setCleaningPlatesInput] = useState<string>('')
  const [cleaningGlassesInput, setCleaningGlassesInput] = useState<string>('')
  const [cleaningPlateRateInput, setCleaningPlateRateInput] = useState<string>('')
  const [cleaningGlassRateInput, setCleaningGlassRateInput] = useState<string>('')
  const [editingRental, setEditingRental] = useState<Rental | null>(null)
  const [editingSalary, setEditingSalary] = useState<SalaryRecord | null>(null)
  const [editingExpense, setEditingExpense] = useState<ExpenseRecord | null>(null)
  const [showLicense, setShowLicense] = useState(false)

  // Load data on start
  useEffect(() => {
    window.electronAPI.getData().then((savedData: any) => {
      if (savedData) {
        // Ensure all necessary structures exist
        const sanitizedData: AppData = {
          settings: {
            ...data.settings,
            ...(savedData.settings || {})
          },
          rentals: Array.isArray(savedData.rentals) ? savedData.rentals : [],
          salaries: Array.isArray(savedData.salaries) ? savedData.salaries : [],
          expenses: Array.isArray(savedData.expenses) ? savedData.expenses : [],
          ownerTransactions: Array.isArray(savedData.ownerTransactions) ? savedData.ownerTransactions : [
            ...(Array.isArray(savedData.withdrawals) ? savedData.withdrawals.map((w: any) => ({
              id: w.id,
              date: w.date,
              from: 'BUSINESS',
              to: w.ownerName,
              amount: w.amount,
              notes: w.notes
            })) : [])
          ]
        }
        if (savedData.theme) setTheme(savedData.theme)
        setData(sanitizedData)
      }
      setLoading(false)
    })
  }, [])

  // Auto-save data whenever it changes
  useEffect(() => {
    if (!loading) {
      window.electronAPI.saveData({ ...data, theme })
    }
  }, [data, theme, loading])

  // --- Calculations ---
  const stats = useMemo(() => {
    const rentals = data.rentals || []
    const salaries = data.salaries || []
    const expenses = data.expenses || []
    const settings = data.settings || { plateCleaningPrice: 0, glassCleaningPrice: 0 }

    const totalEarnings = rentals.reduce((sum, r) => sum + (Number(r.total) || 0) + (Number(r.damageCharge) || 0), 0)
    const activeRentals = rentals.filter(r => !r.isReturned).length
    const totalPlatesOut = rentals.filter(r => !r.isReturned).reduce((sum, r) => sum + (Number(r.plateCount) || 0), 0)
    const totalGlassesOut = rentals.filter(r => !r.isReturned).reduce((sum, r) => sum + (Number(r.glassCount) || 0), 0)

    const platesPendingCleaning = rentals.filter(r => r.isReturned && !r.isCleaned).reduce((sum, r) => {
      const pending = (Number(r.plateCount) || 0) - (Number(r.unusedPlateCount) || 0) - (Number(r.platesCleaned) || 0)
      return sum + Math.max(0, pending)
    }, 0)
    const glassesPendingCleaning = rentals.filter(r => r.isReturned && !r.isCleaned).reduce((sum, r) => {
      const pending = (Number(r.glassCount) || 0) - (Number(r.unusedGlassCount) || 0) - (Number(r.glassesCleaned) || 0)
      return sum + Math.max(0, pending)
    }, 0)

    const totalManualCleaningSalaries = salaries.filter(s => s.rentalId === 'MANUAL').reduce((sum, s) => sum + (Number(s.totalSalary) || 0), 0)

    const totalRentalCleaningCosts = rentals.reduce((sum, r) => {
      const plateRate = Number(settings.plateCleaningPrice) || 0
      const glassRate = Number(settings.glassCleaningPrice) || 0
      const expected = ((Number(r.plateCount) || 0) * plateRate) + ((Number(r.glassCount) || 0) * glassRate)
      return sum + (r.cleaningCost !== undefined ? (Number(r.cleaningCost) || 0) : expected)
    }, 0)

    const totalOtherExpenses = expenses.reduce((sum, e) => sum + (Number(e.amount) || 0), 0)
    const totalPerRentalTravelExpenses = rentals.reduce((sum, r) => sum + (Number(r.travelExpense) || 0), 0)

    const totalExpenses = totalManualCleaningSalaries + totalRentalCleaningCosts + totalOtherExpenses + totalPerRentalTravelExpenses
    const netProfit = totalEarnings - totalExpenses

    const totalDamageCollected = rentals.reduce((sum, r) => sum + (Number(r.damageCharge) || 0), 0)

    return {
      totalEarnings, activeRentals,
      totalPlatesOut, totalGlassesOut,
      platesPendingCleaning, glassesPendingCleaning,
        totalExpenses, netProfit, totalDamageCollected,
        totalPlatesLost: expenses.filter(e => e.category === 'Damage/Loss' && e.description.toLowerCase().includes('plate')).reduce((s, e) => {
          const match = e.description.match(/(\d+)\s*plate/i)
          return s + (match ? parseInt(match[1]) : 0)
        }, 0),
        totalGlassesLost: expenses.filter(e => e.category === 'Damage/Loss' && e.description.toLowerCase().includes('glass')).reduce((s, e) => {
          const match = e.description.match(/(\d+)\s*glass/i)
          return s + (match ? parseInt(match[1]) : 0)
        }, 0)
      }
    }, [data.rentals, data.salaries, data.expenses, data.settings])

  // --- Handlers ---
  const addRental = (rental: Omit<Rental, 'id' | 'isReturned'>) => {
    setData(prev => {
      let maxId = 1000
      prev.rentals.forEach(r => {
        if (r.id.startsWith('ML-')) {
          const num = parseInt(r.id.replace('ML-', ''), 10)
          if (!isNaN(num) && num > maxId) maxId = num
        }
      })

      const newId = `ML-${(maxId + 1).toString().padStart(4, '0')}`

      const newRental: Rental = {
        ...rental,
        id: newId,
        isReturned: false
      }

      return {
        ...prev,
        rentals: [newRental, ...prev.rentals]
      }
    })
  }

  const handleInitiateReturn = (id: string) => {
    setReturningRentalId(id)
    setDamageChargeInput('')
    setUnusedPlatesInput('')
    setUnusedGlassesInput('')
  }

  const confirmReturn = (e: React.FormEvent) => {
    e.preventDefault()
    if (!returningRentalId) return
    const damageCharge = Number(damageChargeInput) || 0
    const unusedPlates = Number(unusedPlatesInput) || 0
    const unusedGlasses = Number(unusedGlassesInput) || 0
    const returnDate = new Date().toISOString().split('T')[0]

    setData(prev => ({
      ...prev,
      rentals: prev.rentals.map(r => {
        if (r.id === returningRentalId) {
          const usedPlates = Math.max(0, r.plateCount - unusedPlates)
          const usedGlasses = Math.max(0, r.glassCount - unusedGlasses)
          const total = (usedPlates * r.platePrice) + (usedGlasses * r.glassPrice) - r.discount
          return {
            ...r,
            isReturned: true,
            returnDate,
            damageCharge,
            unusedPlateCount: unusedPlates,
            unusedGlassCount: unusedGlasses,
            total
          }
        }
        return r
      })
    }))
    setReturningRentalId(null)
  }

  const cancelReturn = () => {
    setReturningRentalId(null)
  }

  const handleInitiateCleaning = (id: string) => {
    const rental = data.rentals.find(r => r.id === id)
    if (!rental) return
    setCleaningRentalId(id)
    setCleaningEmployeeSelect('')
    
    const usedPlates = rental.plateCount - (rental.unusedPlateCount || 0)
    const usedGlasses = rental.glassCount - (rental.unusedGlassCount || 0)
    const pendingPlates = Math.max(0, usedPlates - (rental.platesCleaned || 0))
    const pendingGlasses = Math.max(0, usedGlasses - (rental.glassesCleaned || 0))
    
    setCleaningPlatesInput(pendingPlates.toString())
    setCleaningGlassesInput(pendingGlasses.toString())
    setCleaningPlateRateInput(data.settings.plateCleaningPrice.toString())
    setCleaningGlassRateInput(data.settings.glassCleaningPrice.toString())
  }

  const confirmCleaning = (e: React.FormEvent) => {
    e.preventDefault()
    if (!cleaningRentalId) return
    if (!cleaningEmployeeSelect) {
      alert('Please select an employee.')
      return
    }

    const rental = data.rentals.find(r => r.id === cleaningRentalId)
    if (!rental) return

    const platesToClean = Number(cleaningPlatesInput) || 0
    const glassesToClean = Number(cleaningGlassesInput) || 0
    const plateRate = Number(cleaningPlateRateInput) || 0
    const glassRate = Number(cleaningGlassRateInput) || 0
    const totalSalary = (platesToClean * plateRate) + (glassesToClean * glassRate)

    const newSalary: SalaryRecord = {
      id: Date.now().toString(),
      date: new Date().toISOString().split('T')[0],
      employeeName: cleaningEmployeeSelect,
      rentalId: rental.id,
      plateCount: platesToClean,
      glassCount: glassesToClean,
      plateRate,
      glassRate,
      totalSalary
    }

    setData(prev => ({
      ...prev,
      rentals: prev.rentals.map(r => {
        if (r.id === cleaningRentalId) {
          const newPlatesCleaned = (r.platesCleaned || 0) + platesToClean
          const newGlassesCleaned = (r.glassesCleaned || 0) + glassesToClean
          const usedPlates = r.plateCount - (r.unusedPlateCount || 0)
          const usedGlasses = r.glassCount - (r.unusedGlassCount || 0)
          const isFullyCleaned = newPlatesCleaned >= usedPlates && newGlassesCleaned >= usedGlasses
          
          return {
            ...r,
            platesCleaned: newPlatesCleaned,
            glassesCleaned: newGlassesCleaned,
            isCleaned: isFullyCleaned,
            cleaningCost: (r.cleaningCost || 0) + totalSalary
          }
        }
        return r
      }),
      salaries: [newSalary, ...prev.salaries]
    }))
    setCleaningRentalId(null)
  }

  const cancelCleaning = () => {
    setCleaningRentalId(null)
  }

  /* Removed unused addSalary */

  const addExpense = (record: Omit<ExpenseRecord, 'id'>) => {
    setData(prev => ({
      ...prev,
      expenses: [{ ...record, id: Date.now().toString() }, ...(prev.expenses || [])]
    }))
  }

  const markCleaned = (rentalId: string) => {
    if (!cleaningEmployee) {
      alert('Please enter an Employee Name first.')
      return
    }
    const rental = data.rentals.find(r => r.id === rentalId)
    if (!rental) return

    const usedPlates = rental.plateCount - (rental.unusedPlateCount || 0)
    const usedGlasses = rental.glassCount - (rental.unusedGlassCount || 0)
    const pendingPlates = Math.max(0, usedPlates - (rental.platesCleaned || 0))
    const pendingGlasses = Math.max(0, usedGlasses - (rental.glassesCleaned || 0))

    const totalSalary = (pendingPlates * data.settings.plateCleaningPrice) + (pendingGlasses * data.settings.glassCleaningPrice)

    const newSalary: SalaryRecord = {
      id: Date.now().toString(),
      date: new Date().toISOString().split('T')[0],
      employeeName: cleaningEmployee,
      rentalId: rental.id,
      plateCount: pendingPlates,
      glassCount: pendingGlasses,
      plateRate: data.settings.plateCleaningPrice,
      glassRate: data.settings.glassCleaningPrice,
      totalSalary
    }

    setData(prev => ({
      ...prev,
      rentals: prev.rentals.map(r => r.id === rentalId ? { 
        ...r, 
        platesCleaned: (r.platesCleaned || 0) + pendingPlates,
        glassesCleaned: (r.glassesCleaned || 0) + pendingGlasses,
        isCleaned: true, 
        cleaningCost: (r.cleaningCost || 0) + totalSalary 
      } : r),
      salaries: [newSalary, ...prev.salaries]
    }))
  }

  const resetData = () => {
    if (window.confirm("ARE YOU SURE? This will permanently delete ALL rentals, salaries, and expense records!")) {
      setData({
        settings: { ...data.settings },
        rentals: [],
        salaries: [],
        expenses: [],
        ownerTransactions: []
      })
    }
  }

  const deleteRental = (id: string) => {
    if (window.confirm("Delete this rental record permanently?")) {
      setData(prev => ({
        ...prev,
        rentals: prev.rentals.filter(r => r.id !== id)
      }))
    }
  }

  const deleteSalary = (id: string) => {
    if (window.confirm("Delete this salary record permanently?")) {
      setData(prev => ({
        ...prev,
        salaries: prev.salaries.filter(s => s.id !== id)
      }))
    }
  }

  const deleteExpense = (id: string) => {
    if (window.confirm("Delete this expense record permanently?")) {
      setData(prev => ({
        ...prev,
        expenses: prev.expenses.filter(e => e.id !== id)
      }))
    }
  }

  const updateRental = (updated: Rental) => {
    setData(prev => ({
      ...prev,
      rentals: prev.rentals.map(r => r.id === updated.id ? updated : r)
    }))
    setEditingRental(null)
  }

  /* Removed addSalary as it is handled via markCleaned and manualClean */
  const updateSalary = (updated: SalaryRecord) => {
    setData(prev => ({
      ...prev,
      salaries: prev.salaries.map(s => s.id === updated.id ? updated : s)
    }))
    setEditingSalary(null)
  }

  const updateExpense = (updated: ExpenseRecord) => {
    setData(prev => ({
      ...prev,
      expenses: prev.expenses.map(e => e.id === updated.id ? updated : e)
    }))
    setEditingExpense(null)
  }

  const handleManualClean = () => {
    if (!manualEmployee) {
      alert('Please select an Employee for Manual Cleaning.')
      return
    }
    const platesToClean = Number(manualPlates) || 0
    const glassesToClean = Number(manualGlasses) || 0
    if (platesToClean === 0 && glassesToClean === 0) return

    const totalSalary = (platesToClean * data.settings.plateCleaningPrice) + (glassesToClean * data.settings.glassCleaningPrice)

    const newSalary: SalaryRecord = {
      id: Date.now().toString(),
      date: manualCleaningDate,
      employeeName: manualEmployee,
      rentalId: 'MANUAL',
      plateCount: platesToClean,
      glassCount: glassesToClean,
      plateRate: data.settings.plateCleaningPrice,
      glassRate: data.settings.glassCleaningPrice,
      totalSalary
    }

    setData(prev => ({
      ...prev,
      salaries: [newSalary, ...prev.salaries]
    }))
    setManualPlates('')
    setManualGlasses('')
  }

  const handleAddEmployee = () => {
    if (newEmployeeName && !data.settings.employees?.includes(newEmployeeName)) {
      updateSettings({ ...data.settings, employees: [...(data.settings.employees || []), newEmployeeName] })
      setCleaningEmployee(newEmployeeName)
      setNewEmployeeName('')
    }
  }

  const handleAddOwnerTransaction = (tx: Omit<OwnerTransaction, 'id'>) => {
    const newId = `TX-${Date.now()}`
    setData(prev => ({
      ...prev,
      ownerTransactions: [{ ...tx, id: newId }, ...(prev.ownerTransactions || [])]
    }))
  }

  const deleteOwnerTransaction = (id: string) => {
    if (window.confirm('Delete this transaction record?')) {
      setData(prev => ({
        ...prev,
        ownerTransactions: prev.ownerTransactions.filter(t => t.id !== id)
      }))
    }
  }

  const updateSettings = (newSettings: ItemSettings) => {
    setData(prev => ({ ...prev, settings: newSettings }))
  }

  // --- Export & PDF ---
  const generateOwnerReportPDF = (ownerName: string | null) => {
    try {
      const doc = new jsPDF()
      const primaryRed = [227, 30, 36]
      const brandYellow = [255, 209, 0]
      const darkGrey = [88, 89, 91]

      // --- Logo Drawing (Same as invoice) ---
      doc.setFillColor(primaryRed[0], primaryRed[1], primaryRed[2])
      doc.triangle(20, 15, 32, 17, 35, 45, 'F')
      doc.triangle(20, 15, 35, 45, 22, 48, 'F')
      doc.setFillColor(brandYellow[0], brandYellow[1], brandYellow[2])
      doc.triangle(38, 17, 50, 15, 48, 48, 'F')
      doc.triangle(38, 17, 48, 48, 35, 45, 'F')
      doc.setDrawColor(darkGrey[0], darkGrey[1], darkGrey[2])
      doc.setLineWidth(1)
      doc.line(26, 25, 35, 40)
      doc.line(35, 40, 44, 25)
      doc.setFontSize(14)
      doc.setTextColor(0)
      doc.setFont("helvetica", "bold")
      doc.text('Melon Events', 20, 52)
      doc.setFontSize(9)
      doc.setFont("helvetica", "normal")
      doc.setTextColor(100)
      doc.text('Owner Financial Report', 20, 57)

      // --- Header Right ---
      doc.setFontSize(22)
      doc.setTextColor(primaryRed[0], primaryRed[1], primaryRed[2])
      doc.text('FINANCIAL REPORT', 190, 35, { align: 'right' })
      doc.setFontSize(10)
      doc.setTextColor(100)
      doc.text(`Generated: ${new Date().toLocaleDateString('en-IN')}`, 190, 42, { align: 'right' })

      doc.setDrawColor(200)
      doc.line(20, 58, 190, 58)

      const o1 = data.settings.owner1Name || 'ASHRAF M'
      const o2 = data.settings.owner2Name || 'RASHEED M'
      const txs = data.ownerTransactions || []
      const ownerShare = stats.netProfit / 2

      const getOwnerStats = (name: string, other: string) => {
        const investments = txs.filter(t => t.from === name && t.to === 'BUSINESS').reduce((s, t) => s + t.amount, 0)
        const legacyInvestment = name === o1 ? (data.settings.owner1Investment || 0) : (data.settings.owner2Investment || 0)
        const totalInvested = investments + legacyInvestment
        const profitPayouts = txs.filter(t => t.from === 'BUSINESS' && t.to === name && t.subType === 'profit').reduce((s, t) => s + t.amount, 0)
        const investmentReturns = txs.filter(t => t.from === 'BUSINESS' && t.to === name && t.subType === 'investment').reduce((s, t) => s + t.amount, 0)
        const otherWithdrawals = txs.filter(t => t.from === 'BUSINESS' && t.to === name && !t.subType).reduce((s, t) => s + t.amount, 0)
        const transfersOut = txs.filter(t => t.from === name && t.to === other).reduce((s, t) => s + t.amount, 0)
        const transfersIn = txs.filter(t => t.from === other && t.to === name).reduce((s, t) => s + t.amount, 0)
        const remainingInvestment = totalInvested - investmentReturns
        const remainingProfit = ownerShare - profitPayouts - otherWithdrawals - transfersOut + transfersIn
        return { totalInvested, investmentReturns, remainingInvestment, profitShare: ownerShare, profitPayouts, remainingProfit, totalStanding: remainingInvestment + remainingProfit }
      }

      // --- Summary Section ---
      doc.setFontSize(12)
      doc.setTextColor(0)
      doc.text(`Report for: ${ownerName || 'Combined (All Owners)'}`, 20, 70)

      const summaryData = []
      if (!ownerName || ownerName === o1) {
        const s = getOwnerStats(o1, o2)
        summaryData.push([o1, `Rs. ${s.totalInvested.toFixed(2)}`, `Rs. ${s.profitShare.toFixed(2)}`, `Rs. ${(s.investmentReturns + s.profitPayouts).toFixed(2)}`, `Rs. ${s.totalStanding.toFixed(2)}`])
      }
      if (!ownerName || ownerName === o2) {
        const s = getOwnerStats(o2, o1)
        summaryData.push([o2, `Rs. ${s.totalInvested.toFixed(2)}`, `Rs. ${s.profitShare.toFixed(2)}`, `Rs. ${(s.investmentReturns + s.profitPayouts).toFixed(2)}`, `Rs. ${s.totalStanding.toFixed(2)}`])
      }

      autoTable(doc, {
        startY: 75,
        head: [['Owner Name', 'Invested', 'Profit Share', 'Total Paid', 'Net Balance']],
        body: summaryData,
        theme: 'striped',
        headStyles: { fillColor: [50, 50, 50] }
      })

      // --- Detailed Transactions ---
      const detailedTxs = txs.filter(t => !ownerName || t.from === ownerName || t.to === ownerName)
        .sort((a, b) => new Date(b.date).getTime() - new Date(a.date).getTime())

      if (detailedTxs.length > 0) {
        doc.setFontSize(11)
        doc.text('Transaction History', 20, (doc as any).lastAutoTable.finalY + 15)
        
        autoTable(doc, {
          startY: (doc as any).lastAutoTable.finalY + 20,
          head: [['Date', 'From', 'To', 'Type', 'Notes', 'Amount']],
          body: detailedTxs.map(t => [
            t.date,
            t.from,
            t.to,
            t.subType || '-',
            t.notes,
            `Rs. ${t.amount.toFixed(2)}`
          ]),
          theme: 'grid',
          headStyles: { fillColor: primaryRed as [number, number, number] }
        })
      }

      doc.save(`MelonEvents_Report_${ownerName || 'Both'}_${new Date().toISOString().split('T')[0]}.pdf`)
    } catch (err) {
      console.error("PDF Report Error:", err)
      alert("Failed to generate PDF report.")
    }
  }

  const exportToExcel = (type: 'rentals' | 'salaries' | 'expenses') => {
    let items: any[] = []

    if (type === 'rentals') {
      const sortedRentals = [...data.rentals].sort((a, b) => new Date(b.date).getTime() - new Date(a.date).getTime())
      items = sortedRentals.map(r => {
        const custBill = (Number(r.total) || 0) + (Number(r.damageCharge) || 0)
        const expectedCleaning = ((Number(r.plateCount) || 0) * data.settings.plateCleaningPrice) + ((Number(r.glassCount) || 0) * data.settings.glassCleaningPrice)
        const actualCleaning = r.cleaningCost !== undefined ? (Number(r.cleaningCost) || 0) : expectedCleaning
        const ownerProfit = custBill - (Number(r.travelExpense) || 0) - actualCleaning

        return {
          ID: r.id,
          Date: r.date,
          Customer: r.customerName,
          'Plates Rented': r.plateCount,
          'Glasses Rented': r.glassCount,
          'Plate Price': r.platePrice,
          'Glass Price': r.glassPrice,
          Discount: r.discount || 0,
          'Travel Expense': r.travelExpense || 0,
          'Damage Charge': r.damageCharge || 0,
          'Cleaning Charge': actualCleaning,
          'Customer Bill': custBill,
          'Profit (After Cleaning)': ownerProfit,
          Status: r.isReturned ? 'Returned' : 'Active'
        }
      })
    } else if (type === 'salaries') {
      items = [...data.salaries].sort((a, b) => new Date(b.date).getTime() - new Date(a.date).getTime())
    } else {
      items = [...data.expenses].sort((a, b) => new Date(b.date).getTime() - new Date(a.date).getTime())
    }

    const worksheet = XLSX.utils.json_to_sheet(items)
    const workbook = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(workbook, worksheet, type.charAt(0).toUpperCase() + type.slice(1))

    const excelBuffer = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' })
    window.electronAPI.exportExcel({
      fileName: `MelonEvents_${type}_${new Date().toISOString().split('T')[0]}.xlsx`,
      buffer: Array.from(new Uint8Array(excelBuffer))
    })
  }

  const generatePDF = (rental: Rental) => {
    try {
      const doc = new jsPDF()
      const primaryRed = [227, 30, 36] // Brand Red
      const brandYellow = [255, 209, 0] // Brand Yellow
      const darkGrey = [88, 89, 91] // V-shape color

      // --- Logo Drawing (Vector) ---
      // Left Red Shard
      doc.setFillColor(primaryRed[0], primaryRed[1], primaryRed[2])
      doc.setDrawColor(primaryRed[0], primaryRed[1], primaryRed[2])
      doc.triangle(20, 15, 32, 17, 35, 45, 'F')
      doc.triangle(20, 15, 35, 45, 22, 48, 'F')

      // Right Yellow Shard
      doc.setFillColor(brandYellow[0], brandYellow[1], brandYellow[2])
      doc.setDrawColor(brandYellow[0], brandYellow[1], brandYellow[2])
      doc.triangle(38, 17, 50, 15, 48, 48, 'F')
      doc.triangle(38, 17, 48, 48, 35, 45, 'F')

      // V-shape lines
      doc.setDrawColor(darkGrey[0], darkGrey[1], darkGrey[2])
      doc.setLineWidth(1)
      doc.line(26, 25, 35, 40)
      doc.line(35, 40, 44, 25)

      // White highlight
      doc.setDrawColor(255, 255, 255)
      doc.setLineWidth(0.3)
      doc.line(29, 27, 35, 37)
      doc.line(35, 37, 41, 27)

      // Logo Text
      doc.setFontSize(14)
      doc.setTextColor(0, 0, 0)
      doc.setFont("helvetica", "bold")
      doc.text('Melon Events', 20, 52)

      doc.setFontSize(9)
      doc.setFont("helvetica", "normal")
      doc.setTextColor(100)
      doc.text('Contact: +91 9072266505', 20, 57)

      // --- Header Right ---
      doc.setFontSize(26)
      doc.setTextColor(primaryRed[0], primaryRed[1], primaryRed[2])
      doc.text('INVOICE', 190, 35, { align: 'right' })

      doc.setFontSize(10)
      doc.setTextColor(100)
      doc.setFont("helvetica", "normal")
      doc.text('Premium Catering & Event Rentals', 190, 42, { align: 'right' })

      // Separator Line
      doc.setDrawColor(200)
      doc.setLineWidth(0.5)
      doc.line(20, 58, 190, 58)

      // --- Details Section ---
      doc.setFontSize(10)
      doc.setTextColor(80)
      doc.text(`Invoice No: ${rental.id}`, 20, 70)
      doc.text(`Rental Date: ${rental.date}`, 20, 76)
      if (rental.returnDate) {
        doc.text(`Return Date: ${rental.returnDate}`, 20, 82)
      }

      doc.setFont("helvetica", "bold")
      doc.text('BILL TO:', 140, 70)
      doc.setFont("helvetica", "normal")
      doc.setFontSize(12)
      doc.setTextColor(0)
      doc.text(rental.customerName, 140, 78)

      // --- Items Table ---
      const tableData = [['Description', 'Qty', 'Rate', 'Total']]
      
      tableData.push(['Fiber Plates Rental', rental.plateCount.toString(), `Rs. ${rental.platePrice.toFixed(2)}`, `Rs. ${(rental.plateCount * rental.platePrice).toFixed(2)}`])
      if (rental.unusedPlateCount && rental.unusedPlateCount > 0) {
        tableData.push(['Unused Plates (Credit)', `-${rental.unusedPlateCount}`, `Rs. ${rental.platePrice.toFixed(2)}`, `-Rs. ${(rental.unusedPlateCount * rental.platePrice).toFixed(2)}`])
      }
      
      tableData.push(['Glasses Rental', rental.glassCount.toString(), `Rs. ${rental.glassPrice.toFixed(2)}`, `Rs. ${(rental.glassCount * rental.glassPrice).toFixed(2)}`])
      if (rental.unusedGlassCount && rental.unusedGlassCount > 0) {
        tableData.push(['Unused Glasses (Credit)', `-${rental.unusedGlassCount}`, `Rs. ${rental.glassPrice.toFixed(2)}`, `-Rs. ${(rental.unusedGlassCount * rental.glassPrice).toFixed(2)}`])
      }

      if (rental.damageCharge && rental.damageCharge > 0) {
        tableData.push(['Damage / Missing Charges', '-', '-', `Rs. ${rental.damageCharge.toFixed(2)}`])
      }

      autoTable(doc, {
        startY: 95,
        head: [tableData[0]],
        body: tableData.slice(1),
        theme: 'grid',
        headStyles: { fillColor: primaryRed as [number, number, number], textColor: 255, fontStyle: 'bold' },
        styles: { fontSize: 10, cellPadding: 6, valign: 'middle' },
        columnStyles: {
          1: { halign: 'center', cellWidth: 20 },
          2: { halign: 'right', cellWidth: 35 },
          3: { halign: 'right', cellWidth: 35 }
        }
      })

      const finalY = (doc as any).lastAutoTable.finalY + 15
      const subtotal = ((rental.plateCount - (rental.unusedPlateCount || 0)) * rental.platePrice) + 
                       ((rental.glassCount - (rental.unusedGlassCount || 0)) * rental.glassPrice)
      const totalWithDamage = subtotal - (rental.discount || 0) + (rental.damageCharge || 0)

      // --- Totals Section ---
      doc.setFontSize(10)
      doc.setTextColor(100)
      doc.text('Subtotal:', 140, finalY)
      doc.text(`Rs. ${subtotal.toFixed(2)}`, 190, finalY, { align: 'right' })

      if (rental.discount > 0) {
        doc.text('Discount:', 140, finalY + 8)
        doc.setTextColor(primaryRed[0], primaryRed[1], primaryRed[2])
        doc.text(`Rs. ${rental.discount.toFixed(2)}`, 190, finalY + 8, { align: 'right' })
        doc.setTextColor(100)
      }

      if (rental.damageCharge && rental.damageCharge > 0) {
        doc.text('Damage Charges:', 140, finalY + 16)
        doc.text(`Rs. ${rental.damageCharge.toFixed(2)}`, 190, finalY + 16, { align: 'right' })
      }

      // Grand Total Highlight
      doc.setFillColor(primaryRed[0], primaryRed[1], primaryRed[2])
      doc.rect(115, finalY + 22, 80, 14, 'F')
      doc.setFontSize(11)
      doc.setFont("helvetica", "bold")
      doc.setTextColor(255, 255, 255)
      doc.text('GRAND TOTAL :', 120, finalY + 31)
      doc.setFontSize(14)
      doc.text(`Rs. ${totalWithDamage.toFixed(2)}`, 190, finalY + 31, { align: 'right' })

      // --- Footer ---
      doc.setFontSize(10)
      doc.setFont("helvetica", "italic")
      doc.setTextColor(150)
      doc.text('Thank you for choosing Melon Events!', 105, 275, { align: 'center' })
      doc.text('Quality Service for Your Special Moments', 105, 282, { align: 'center' })

      const pdfOutput = doc.output('arraybuffer')
      window.electronAPI.savePDF({
        fileName: `Invoice_${rental.customerName}_${rental.id}.pdf`,
        buffer: Array.from(new Uint8Array(pdfOutput))
      })
    } catch (error) {
      console.error("PDF generation failed:", error)
      alert("Error generating PDF invoice. Please check the console.")
    }
  }

  if (loading) return (
    <div style={{
      height: '100vh', width: '100vw', display: 'flex', flexDirection: 'column',
      alignItems: 'center', justifyContent: 'center', background: '#0f172a', color: 'white'
    }}>
      <img src="logo.png" alt="Melon Events" style={{ width: '120px', height: '120px', marginBottom: '20px' }} />
      <div style={{ fontSize: '1.5rem', fontWeight: 'bold' }}>Melon Events</div>
      <div style={{ marginTop: '10px', color: '#94a3b8' }}>Loading your dashboard...</div>
    </div>
  )

  return (
    <div className={theme === 'light' ? 'light-theme' : ''} style={{ display: 'flex', height: '100vh', width: '100vw' }}>
      {/* Sidebar */}
      <nav className="sidebar">
        <div className="logo" style={{ marginBottom: '2.5rem', display: 'flex', alignItems: 'center', gap: '14px', padding: '0 0.5rem' }}>
          <img src="logo-mark.png" alt="Logo" style={{ width: '52px', height: '52px', objectFit: 'contain' }} />
          <span style={{ color: 'var(--text-primary)', fontWeight: '900', fontSize: '1.4rem', letterSpacing: '-0.03em', lineHeight: '1' }}>Melon Events</span>
        </div>
        <div className="nav-links">
          <div
            className={`nav-item ${activeTab === 'dashboard' ? 'active' : ''}`}
            onClick={() => setActiveTab('dashboard')}
          >
            <LayoutDashboard size={20} /> Dashboard
          </div>
          <div
            className={`nav-item ${activeTab === 'rentals' ? 'active' : ''}`}
            onClick={() => setActiveTab('rentals')}
          >
            <Receipt size={20} /> Rentals & Billing
          </div>
          <div
            className={`nav-item ${activeTab === 'salaries' ? 'active' : ''}`}
            onClick={() => setActiveTab('salaries')}
          >
            <Briefcase size={20} /> Cleaning Salary
          </div>
          <div
            className={`nav-item ${activeTab === 'expenses' ? 'active' : ''}`}
            onClick={() => setActiveTab('expenses')}
          >
            <Wallet size={20} /> Business Expenses
          </div>
          <div
            className={`nav-item ${activeTab === 'inventory' ? 'active' : ''}`}
            onClick={() => setActiveTab('inventory')}
          >
            <Settings size={20} /> Settings
          </div>
          <div
            className={`nav-item ${activeTab === 'owner-profit' ? 'active' : ''}`}
            onClick={() => setActiveTab('owner-profit')}
          >
            <IndianRupee size={20} /> Owner Profit
          </div>
          <div
            className={`nav-item ${activeTab === 'about' ? 'active' : ''}`}
            onClick={() => setActiveTab('about')}
          >
            <Info size={20} /> About
          </div>

          <div
            className="nav-item"
            onClick={() => setTheme(theme === 'dark' ? 'light' : 'dark')}
            style={{ marginTop: '1rem', border: '1px solid var(--border-color)' }}
          >
            {theme === 'dark' ? <Sun size={20} /> : <Moon size={20} />}
            {theme === 'dark' ? 'Light Mode' : 'Dark Mode'}
          </div>
        </div>

        <div style={{
          marginTop: 'auto',
          padding: '1.5rem',
          borderTop: '1px solid var(--border-color)',
          fontSize: '0.7rem',
          color: 'var(--text-secondary)',
          lineHeight: '1.4'
        }}>
          <div>© 2026 Melon Events</div>
          <div style={{ color: 'var(--text-primary)', fontWeight: '500' }}>Developed by Alinshan</div>
        </div>
      </nav>

      {/* Main Content */}
      <main className="main-content">
        <AnimatePresence mode="wait">
          {activeTab === 'dashboard' && (
            <motion.div
              key="dash"
              initial={{ opacity: 0, y: 20 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, y: -20 }}
            >
              <h1>Dashboard Overview</h1>
              <div style={{ display: 'grid', gridTemplateColumns: '1fr 320px', gap: '1.5rem', alignItems: 'start' }}>
                {/* Left Column */}
                <div>
                  <div className="stats-grid" style={{ gridTemplateColumns: 'repeat(5, 1fr)', marginBottom: '1.5rem' }}>
                    <div className="stat-card">
                      <span className="stat-label">Total Earnings</span>
                      <span className="stat-value">₹{stats.totalEarnings.toFixed(2)}</span>
                      <IndianRupee size={24} style={{ color: 'var(--success)', marginTop: 'auto' }} />
                    </div>
                    <div className="stat-card">
                      <span className="stat-label">Total Expenses</span>
                      <span className="stat-value">₹{stats.totalExpenses.toFixed(2)}</span>
                      <Wallet size={24} style={{ color: 'var(--danger)', marginTop: 'auto' }} />
                    </div>
                    <div className="stat-card">
                      <span className="stat-label">Net Profit</span>
                      <span className="stat-value" style={{ color: stats.netProfit >= 0 ? 'var(--success)' : 'var(--danger)' }}>
                        ₹{stats.netProfit.toFixed(2)}
                      </span>
                      <Receipt size={24} style={{ color: stats.netProfit >= 0 ? 'var(--success)' : 'var(--danger)', marginTop: 'auto' }} />
                    </div>
                    <div className="stat-card">
                      <span className="stat-label">Damage Coll.</span>
                      <span className="stat-value">₹{stats.totalDamageCollected.toFixed(2)}</span>
                      <AlertTriangle size={24} style={{ color: 'var(--warning)', marginTop: 'auto' }} />
                    </div>
                    <div className="stat-card">
                      <span className="stat-label">Active Rentals</span>
                      <span className="stat-value">{stats.activeRentals}</span>
                      <Calendar size={24} style={{ color: 'var(--accent-color)', marginTop: 'auto' }} />
                    </div>
                  </div>

                  <div className="card" style={{ marginBottom: '1.5rem' }}>
                    <div style={{ display: 'flex', alignItems: 'center', marginBottom: '1rem' }}>
                      <AlertTriangle size={20} style={{ marginRight: '0.5rem', color: 'var(--danger)' }} />
                      <h3 style={{ margin: 0 }}>Report Damage / Loss</h3>
                    </div>
                    <p style={{ fontSize: '0.85rem', color: 'var(--text-secondary)', marginBottom: '1rem' }}>
                      Log items broken or missing after cleaning that cannot be charged to the customer.
                    </p>
                    <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr 1fr', gap: '1rem' }}>
                      <button 
                        className="btn-secondary" 
                        style={{ borderColor: 'var(--danger)', color: 'var(--danger)' }}
                        onClick={() => {
                          const count = prompt("How many PLATES were lost/broken?")
                          if (count && !isNaN(Number(count))) {
                            const amount = Number(count) * (data.settings.plateLossPrice || 0)
                            addExpense({
                              date: new Date().toISOString().split('T')[0],
                              category: 'Damage/Loss',
                              description: `${count} Plates lost/broken`,
                              amount
                            })
                          }
                        }}
                      >
                        Loss: Plates
                      </button>
                      <button 
                        className="btn-secondary" 
                        style={{ borderColor: 'var(--danger)', color: 'var(--danger)' }}
                        onClick={() => {
                          const count = prompt("How many GLASSES were lost/broken?")
                          if (count && !isNaN(Number(count))) {
                            const amount = Number(count) * (data.settings.glassLossPrice || 0)
                            addExpense({
                              date: new Date().toISOString().split('T')[0],
                              category: 'Damage/Loss',
                              description: `${count} Glasses lost/broken`,
                              amount
                            })
                          }
                        }}
                      >
                        Loss: Glasses
                      </button>
                      <button 
                        className="btn-secondary" 
                        onClick={() => {
                          const desc = prompt("Describe the damage (e.g., Damaged Carry Box):")
                          const amt = prompt("Estimated loss amount (₹):")
                          if (desc && amt && !isNaN(Number(amt))) {
                            addExpense({
                              date: new Date().toISOString().split('T')[0],
                              category: 'Damage/Loss',
                              description: desc,
                              amount: Number(amt)
                            })
                          }
                        }}
                      >
                        Other Loss
                      </button>
                    </div>
                  </div>

                  <div className="card">
                    <h3>Inventory Status</h3>
                    <div style={{ marginTop: '1rem' }}>
                      <div style={{ display: 'flex', justifyContent: 'space-between', marginBottom: '0.5rem' }}>
                        <span style={{ color: 'var(--text-secondary)', display: 'flex', alignItems: 'center', gap: '5px' }}>
                          <Package size={16} /> Plates in Stock
                        </span>
                        <span style={{ fontWeight: '700' }}>{Math.max(0, data.settings.totalPlates - stats.totalPlatesOut - stats.totalPlatesLost)} / {data.settings.totalPlates}</span>
                      </div>
                      <div className="progress-bg" style={{ height: '8px', marginBottom: '1rem' }}>
                        <div className="progress-fill" style={{ 
                          width: `${((data.settings.totalPlates - stats.totalPlatesOut - stats.totalPlatesLost) / data.settings.totalPlates) * 100}%`,
                          backgroundColor: 'var(--success)'
                        }}></div>
                      </div>
                      <div style={{ display: 'flex', justifyContent: 'space-between', marginBottom: '0.5rem' }}>
                        <span style={{ color: 'var(--text-secondary)', display: 'flex', alignItems: 'center', gap: '5px' }}>
                          <Wine size={16} /> Glasses in Stock
                        </span>
                        <span style={{ fontWeight: '700' }}>{Math.max(0, data.settings.totalGlasses - stats.totalGlassesOut - stats.totalGlassesLost)} / {data.settings.totalGlasses}</span>
                      </div>
                      <div className="progress-bg" style={{ height: '8px' }}>
                        <div className="progress-fill" style={{ 
                          width: `${((data.settings.totalGlasses - stats.totalGlassesOut - stats.totalGlassesLost) / data.settings.totalGlasses) * 100}%`,
                          backgroundColor: 'var(--accent-color)'
                        }}></div>
                      </div>
                      { (stats.totalPlatesLost > 0 || stats.totalGlassesLost > 0) && (
                        <div style={{ marginTop: '1rem', fontSize: '0.8rem', color: 'var(--danger)', fontStyle: 'italic' }}>
                          * Note: {stats.totalPlatesLost} plates and {stats.totalGlassesLost} glasses have been recorded as permanent loss.
                        </div>
                      )}
                    </div>
                  </div>
                </div>

                {/* Right Column — Booking Calendar */}
                <BookingCalendar rentals={data.rentals} />
              </div>
            </motion.div>
          )}


          {activeTab === 'rentals' && (
            <motion.div
              key="rentals"
              initial={{ opacity: 0, y: 20 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, y: -20 }}
            >
              <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                <h1>Rentals & Billing</h1>
                <button className="btn-secondary" onClick={() => exportToExcel('rentals')}>
                  <Download size={18} /> Export Excel
                </button>
              </div>

              <NewRentalForm
                settings={data.settings}
                onSubmit={addRental}
              />

              <div className="card" style={{ padding: '0' }}>
                <table>
                  <thead>
                    <tr>
                      <th>Rental Date</th>
                      <th>Customer</th>
                      <th>Items (P/G)</th>
                      <th>Status</th>
                      <th>Return Date</th>
                      <th>Cust. Bill</th>
                      <th>Owner Profit</th>
                      <th>Actions</th>
                    </tr>
                  </thead>
                  <tbody>
                    {[...data.rentals].sort((a, b) => new Date(b.date).getTime() - new Date(a.date).getTime()).map(rental => {
                      const custBill = rental.total + (rental.damageCharge || 0)
                      const expectedCleaning = (rental.plateCount * data.settings.plateCleaningPrice) + (rental.glassCount * data.settings.glassCleaningPrice)
                      const actualCleaning = rental.cleaningCost !== undefined ? rental.cleaningCost : expectedCleaning
                      const ownerProfit = custBill - (rental.travelExpense || 0) - actualCleaning

                      return (
                        <tr key={rental.id}>
                          <td>{rental.date}</td>
                          <td>{rental.customerName}</td>
                          <td>
                            <div>{rental.plateCount} / {rental.glassCount}</div>
                            {(rental.unusedPlateCount || rental.unusedGlassCount) ? (
                              <div style={{ fontSize: '0.7rem', color: 'var(--text-secondary)' }}>
                                Unused: {rental.unusedPlateCount || 0}P / {rental.unusedGlassCount || 0}G
                              </div>
                            ) : null}
                          </td>
                          <td style={{ fontWeight: '600' }}>₹{custBill.toFixed(2)}</td>
                          <td>
                            <span className={`badge ${rental.isReturned ? 'badge-success' : 'badge-warning'}`}>
                              {rental.isReturned ? 'Returned' : 'Active'}
                            </span>
                          </td>
                          <td>{rental.returnDate || '-'}</td>
                          <td style={{ fontWeight: '600', color: ownerProfit >= 0 ? 'var(--success)' : 'var(--danger)' }}>
                            ₹{ownerProfit.toFixed(2)}
                          </td>
                          <td style={{ display: 'flex', gap: '0.5rem' }}>
                            {!rental.isReturned && (
                              <button className="btn-secondary" title="Mark Returned" style={{ padding: '0.5rem' }} onClick={() => handleInitiateReturn(rental.id)}>
                                <CheckCircle size={16} />
                              </button>
                            )}
                            {rental.isReturned && !rental.isCleaned && (
                              <button className="btn-secondary" title="Send to Cleaning" style={{ padding: '0.5rem', color: 'var(--accent-color)' }} onClick={() => handleInitiateCleaning(rental.id)}>
                                <Droplets size={16} />
                              </button>
                            )}
                            <button className="btn-secondary" title="Print Invoice" style={{ padding: '0.5rem' }} onClick={() => generatePDF(rental)}>
                              <Printer size={16} />
                            </button>
                            <button className="btn-secondary" title="Edit Rental" style={{ padding: '0.5rem' }} onClick={() => setEditingRental(rental)}>
                              <Edit2 size={16} />
                            </button>
                            <button className="btn-secondary" title="Delete Rental" style={{ padding: '0.5rem', color: 'var(--danger)' }} onClick={() => deleteRental(rental.id)}>
                              <Trash2 size={16} />
                            </button>
                          </td>
                        </tr>
                      )
                    })}
                  </tbody>
                </table>
              </div>
            </motion.div>
          )}

          {activeTab === 'salaries' && (
            <motion.div
              key="salaries"
              initial={{ opacity: 0, y: 20 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, y: -20 }}
            >
              <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                <h1>Cleaning Salaries</h1>
                <button className="btn-secondary" onClick={() => exportToExcel('salaries')}>
                  <Download size={18} /> Export Excel
                </button>
              </div>

              <div className="card" style={{ marginBottom: '1.5rem' }}>
                <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '1rem' }}>
                  <h3>Pending Cleaning</h3>
                  <div className="form-group" style={{ display: 'flex', alignItems: 'center', gap: '0.5rem', margin: 0 }}>
                    <label style={{ margin: 0 }}>Assign To Employee:</label>
                    <select
                      value={cleaningEmployee}
                      onChange={e => setCleaningEmployee(e.target.value)}
                      style={{ padding: '0.4rem', width: '150px' }}
                    >
                      <option value="">Select...</option>
                      {data.settings.employees?.map(emp => (
                        <option key={emp} value={emp}>{emp}</option>
                      ))}
                    </select>
                    <div style={{ display: 'flex', gap: '0.5rem' }}>
                      <input
                        type="text"
                        placeholder="New name"
                        value={newEmployeeName}
                        onChange={e => setNewEmployeeName(e.target.value)}
                        style={{ padding: '0.4rem', width: '120px' }}
                      />
                      <button className="btn-secondary" onClick={handleAddEmployee} style={{ padding: '0.4rem 0.8rem' }}>
                        Add
                      </button>
                    </div>
                  </div>
                </div>

                <table style={{ background: 'transparent' }}>
                  <thead>
                    <tr>
                      <th>Date</th>
                      <th>Customer</th>
                      <th>Items (P/G)</th>
                      <th>Cleaning Fee</th>
                      <th>Action</th>
                    </tr>
                  </thead>
                  <tbody>
                    {data.rentals.filter(r => r.isReturned && !r.isCleaned).length === 0 ? (
                      <tr>
                        <td colSpan={5} style={{ textAlign: 'center', color: 'var(--text-secondary)', padding: '2rem' }}>
                          No pending cleaning items. All caught up!
                        </td>
                      </tr>
                    ) : (
                      [...data.rentals].filter(r => r.isReturned && !r.isCleaned).sort((a, b) => new Date(b.date).getTime() - new Date(a.date).getTime()).map(rental => (
                        <tr key={rental.id}>
                          <td>{rental.date}</td>
                          <td>{rental.customerName}</td>
                          <td>
                            {Math.max(0, rental.plateCount - (rental.unusedPlateCount || 0) - (rental.platesCleaned || 0))} / {Math.max(0, rental.glassCount - (rental.unusedGlassCount || 0) - (rental.glassesCleaned || 0))}
                          </td>
                          <td style={{ color: 'var(--success)', fontWeight: '600' }}>
                            ₹{((Math.max(0, rental.plateCount - (rental.unusedPlateCount || 0) - (rental.platesCleaned || 0)) * data.settings.plateCleaningPrice) + (Math.max(0, rental.glassCount - (rental.unusedGlassCount || 0) - (rental.glassesCleaned || 0)) * data.settings.glassCleaningPrice)).toFixed(2)}
                          </td>
                          <td>
                            <button
                              className="btn-primary"
                              style={{ padding: '0.4rem 0.8rem', background: 'var(--success)' }}
                              onClick={() => markCleaned(rental.id)}
                            >
                              <CheckCircle size={14} style={{ marginRight: '4px' }} /> Mark Cleaned
                            </button>
                          </td>
                        </tr>
                      ))
                    )}

                    {/* Manual Cleaning Row */}
                    <tr style={{ borderTop: '2px dashed var(--border-color)', background: 'rgba(255,255,255,0.02)' }}>
                      <td style={{ fontStyle: 'italic', color: 'var(--text-secondary)' }}>
                        <input
                          type="date"
                          value={manualCleaningDate}
                          onChange={e => setManualCleaningDate(e.target.value)}
                          style={{ padding: '0.2rem', fontSize: '0.8rem', background: 'transparent', border: '1px solid var(--border-color)', color: 'var(--text-primary)' }}
                        />
                      </td>
                      <td>
                        <div style={{ fontWeight: '600', color: 'var(--accent-color)', marginBottom: '0.3rem' }}>Manual / Dusted</div>
                        <select
                          value={manualEmployee}
                          onChange={e => setManualEmployee(e.target.value)}
                          style={{ padding: '0.2rem', width: '130px', fontSize: '0.8rem' }}
                        >
                          <option value="">Select Employee</option>
                          {data.settings.employees?.map(emp => (
                            <option key={emp} value={emp}>{emp}</option>
                          ))}
                        </select>
                      </td>
                      <td style={{ display: 'flex', gap: '0.5rem', alignItems: 'center' }}>
                        <input
                          type="number"
                          min="0"
                          value={manualPlates}
                          onChange={e => setManualPlates(e.target.value)}
                          style={{ width: '60px', padding: '0.2rem 0.4rem' }}
                        /> P /
                        <input
                          type="number"
                          min="0"
                          value={manualGlasses}
                          onChange={e => setManualGlasses(e.target.value)}
                          style={{ width: '60px', padding: '0.2rem 0.4rem' }}
                        /> G
                      </td>
                      <td style={{ color: 'var(--success)', fontWeight: '600' }}>
                        ₹{((Number(manualPlates) * data.settings.plateCleaningPrice) + (Number(manualGlasses) * data.settings.glassCleaningPrice)).toFixed(2)}
                      </td>
                      <td>
                        <button
                          className="btn-secondary"
                          style={{ padding: '0.4rem 0.8rem', borderColor: 'var(--accent-color)', color: 'var(--accent-color)' }}
                          onClick={handleManualClean}
                          disabled={!manualPlates && !manualGlasses}
                        >
                          <CheckCircle size={14} style={{ marginRight: '4px' }} /> Mark Cleaned
                        </button>
                      </td>
                    </tr>
                  </tbody>
                </table>
              </div>

              <div className="card" style={{ padding: '0' }}>
                <table>
                  <thead>
                    <tr>
                      <th>Date</th>
                      <th>Employee</th>
                      <th>Counts (P/G)</th>
                      <th>Rate (P/G)</th>
                      <th>Total Salary</th>
                      <th>Actions</th>
                    </tr>
                  </thead>
                  <tbody>
                    {data.salaries.map(salary => (
                      <tr key={salary.id}>
                        <td>{salary.date}</td>
                        <td>{salary.employeeName}</td>
                        <td>{salary.plateCount} / {salary.glassCount}</td>
                        <td>{salary.plateRate} / {salary.glassRate}</td>
                        <td style={{ fontWeight: '600', color: 'var(--success)' }}>₹{salary.totalSalary}</td>
                        <td style={{ display: 'flex', gap: '0.5rem' }}>
                          <button className="btn-secondary" title="Edit Salary" style={{ padding: '0.4rem' }} onClick={() => setEditingSalary(salary)}>
                            <Edit2 size={14} />
                          </button>
                          <button className="btn-secondary" title="Delete Salary" style={{ padding: '0.4rem', color: 'var(--danger)' }} onClick={() => deleteSalary(salary.id)}>
                            <Trash2 size={14} />
                          </button>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </motion.div>
          )}

          {activeTab === 'expenses' && (
            <motion.div
              key="expenses"
              initial={{ opacity: 0, y: 20 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, y: -20 }}
            >
              <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                <h1>Business Expenses</h1>
                <button className="btn-secondary" onClick={() => exportToExcel('expenses')}>
                  <Download size={18} /> Export Excel
                </button>
              </div>

              <ExpenseForm onSubmit={addExpense} />

              <div className="card" style={{ padding: '0' }}>
                <table>
                  <thead>
                    <tr>
                      <th>Date</th>
                      <th>Category</th>
                      <th>Description</th>
                      <th>Amount</th>
                      <th>Actions</th>
                    </tr>
                  </thead>
                  <tbody>
                    {data.expenses?.length === 0 ? (
                      <tr>
                        <td colSpan={5} style={{ textAlign: 'center', color: 'var(--text-secondary)', padding: '2rem' }}>
                          No expenses recorded yet.
                        </td>
                      </tr>
                    ) : (
                      data.expenses?.map(exp => (
                        <tr key={exp.id}>
                          <td>{exp.date}</td>
                          <td>
                            <span style={{
                              background: 'var(--border-color)',
                              padding: '0.2rem 0.6rem',
                              borderRadius: '12px',
                              fontSize: '0.8rem',
                              color: 'var(--text-primary)'
                            }}>
                              {exp.category}
                            </span>
                          </td>
                          <td>{exp.description}</td>
                          <td style={{ color: 'var(--danger)', fontWeight: '600' }}>₹{exp.amount.toFixed(2)}</td>
                          <td style={{ display: 'flex', gap: '0.5rem' }}>
                            <button className="btn-secondary" title="Edit Expense" style={{ padding: '0.4rem' }} onClick={() => setEditingExpense(exp)}>
                              <Edit2 size={14} />
                            </button>
                            <button className="btn-secondary" title="Delete Expense" style={{ padding: '0.4rem', color: 'var(--danger)' }} onClick={() => deleteExpense(exp.id)}>
                              <Trash2 size={14} />
                            </button>
                          </td>
                        </tr>
                      ))
                    )}
                  </tbody>
                </table>
              </div>
            </motion.div>
          )}

          {activeTab === 'inventory' && (
            <motion.div
              key="settings"
              initial={{ opacity: 0, y: 20 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, y: -20 }}
            >
              <h1>System Settings</h1>
              <div className="card" style={{ maxWidth: '500px' }}>
                <h3>Update Seasonal Pricing & Stock</h3>
                <p style={{ color: 'var(--text-secondary)', fontSize: '0.875rem', marginBottom: '1.5rem' }}>
                  These prices will be applied to all new rentals and cleaning records.
                </p>
                <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '1rem' }}>
                  <div className="form-group">
                    <label>Total Plates Stock</label>
                    <input
                      type="number"
                      value={data.settings.totalPlates || 0}
                      onChange={e => updateSettings({ ...data.settings, totalPlates: Number(e.target.value) })}
                    />
                  </div>
                  <div className="form-group">
                    <label>Total Glasses Stock</label>
                    <input
                      type="number"
                      value={data.settings.totalGlasses || 0}
                      onChange={e => updateSettings({ ...data.settings, totalGlasses: Number(e.target.value) })}
                    />
                  </div>
                </div>
                <hr style={{ margin: '2rem 0', borderColor: 'var(--border-color)' }} />
                <h3>Inventory Loss Prices (Internal)</h3>
                <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '1rem', marginTop: '1rem' }}>
                  <div className="form-group">
                    <label>Plate Loss Cost (₹)</label>
                    <input
                      type="number"
                      value={data.settings.plateLossPrice || 0}
                      onChange={e => updateSettings({ ...data.settings, plateLossPrice: Number(e.target.value) })}
                    />
                    <p style={{ fontSize: '0.7rem', color: 'var(--text-secondary)' }}>Cost to business when 1 plate is lost/broken.</p>
                  </div>
                  <div className="form-group">
                    <label>Glass Loss Cost (₹)</label>
                    <input
                      type="number"
                      value={data.settings.glassLossPrice || 0}
                      onChange={e => updateSettings({ ...data.settings, glassLossPrice: Number(e.target.value) })}
                    />
                    <p style={{ fontSize: '0.7rem', color: 'var(--text-secondary)' }}>Cost to business when 1 glass is lost/broken.</p>
                  </div>
                </div>
                <hr style={{ margin: '2rem 0', borderColor: 'var(--border-color)' }} />
                <div className="form-group">
                  <label>Plate Rent Price (₹)</label>
                  <input
                    type="number"
                    step="0.01"
                    value={data.settings.platePrice}
                    onChange={e => updateSettings({ ...data.settings, platePrice: Number(e.target.value) })}
                  />
                </div>
                <div className="form-group">
                  <label>Glass Rent Price (₹)</label>
                  <input
                    type="number"
                    step="0.01"
                    value={data.settings.glassPrice}
                    onChange={e => updateSettings({ ...data.settings, glassPrice: Number(e.target.value) })}
                  />
                </div>
                <hr style={{ margin: '2rem 0', borderColor: 'var(--border-color)' }} />
                <div className="form-group">
                  <label>Plate Cleaning Salary (₹ per plate)</label>
                  <input
                    type="number"
                    step="0.01"
                    value={data.settings.plateCleaningPrice}
                    onChange={e => updateSettings({ ...data.settings, plateCleaningPrice: Number(e.target.value) })}
                  />
                </div>
                <div className="form-group">
                  <label>Glass Cleaning Salary (₹ per glass)</label>
                  <input
                    type="number"
                    step="0.01"
                    value={data.settings.glassCleaningPrice}
                    onChange={e => updateSettings({ ...data.settings, glassCleaningPrice: Number(e.target.value) })}
                  />
                </div>
                <div style={{ marginTop: '1.5rem', display: 'flex', justifyContent: 'flex-start' }}>
                  <button className="btn-secondary" onClick={() => {
                    if (window.confirm("Restore default prices and cleaning charges? This will not affect your existing rental records.")) {
                      updateSettings({
                        ...data.settings,
                        platePrice: 4,
                        glassPrice: 1,
                        plateCleaningPrice: 1,
                        glassCleaningPrice: 0.4
                      })
                    }
                  }}>
                    <RotateCcw size={16} style={{ marginRight: '0.5rem' }} /> Restore Default Prices
                  </button>
                </div>
                <hr style={{ margin: '2rem 0', borderColor: 'var(--border-color)' }} />
                <div style={{ marginTop: '2rem' }}>
                  <h3 style={{ color: 'var(--danger)' }}>Danger Zone</h3>
                  <p style={{ color: 'var(--text-secondary)', fontSize: '0.875rem', marginBottom: '1rem' }}>
                    Permanently delete all rental history, salary records, and business expenses. This cannot be undone.
                  </p>
                  <button className="btn-secondary" style={{ color: 'var(--danger)', borderColor: 'var(--danger)' }} onClick={resetData}>
                    <Trash2 size={18} style={{ marginRight: '0.5rem' }} /> Clear All Records
                  </button>
                </div>
              </div>
            </motion.div>
          )}

          {activeTab === 'owner-profit' && (
            <motion.div
              key="owner-profit"
              initial={{ opacity: 0, y: 20 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, y: -20 }}
            >
              <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '2rem' }}>
                <h1>Owner Profit</h1>
                <div style={{ display: 'flex', gap: '0.5rem' }}>
                  <button className="btn-secondary" onClick={() => generateOwnerReportPDF(null)}>
                    <Download size={18} /> Combined PDF
                  </button>
                  <button className="btn-secondary" onClick={() => generateOwnerReportPDF(data.settings.owner1Name || 'ASHRAF M')}>
                    <Download size={18} /> {data.settings.owner1Name?.split(' ')[0] || 'ASHRAF'} PDF
                  </button>
                  <button className="btn-secondary" onClick={() => generateOwnerReportPDF(data.settings.owner2Name || 'RASHEED M')}>
                    <Download size={18} /> {data.settings.owner2Name?.split(' ')[0] || 'RASHEED'} PDF
                  </button>
                </div>
              </div>

              {(() => {
                const rentals = data.rentals || []
                const totalRevenue = rentals.reduce((sum, r) => sum + (Number(r.total) || 0) + (Number(r.damageCharge) || 0), 0)
                const totalTravelExpenses = rentals.reduce((sum, r) => sum + (Number(r.travelExpense) || 0), 0)
                const totalCleaningCosts = (data.salaries || []).reduce((sum, s) => sum + (Number(s.totalSalary) || 0), 0)
                const totalOtherExpenses = (data.expenses || []).reduce((sum, e) => sum + (Number(e.amount) || 0), 0)
                const netProfit = totalRevenue - totalTravelExpenses - totalCleaningCosts - totalOtherExpenses
                const ownerShare = netProfit / 2

                return (
                  <>
                    <div className="dashboard-grid" style={{ marginBottom: '2rem' }}>
                      <div className="stat-card">
                        <div className="stat-title">Total Business Revenue</div>
                        <div className="stat-value" style={{ color: 'var(--success)' }}>₹{totalRevenue.toFixed(2)}</div>
                      </div>
                      <div className="stat-card">
                        <div className="stat-title">Net Profit</div>
                        <div className="stat-value" style={{ color: netProfit >= 0 ? 'var(--success)' : 'var(--danger)' }}>₹{netProfit.toFixed(2)}</div>
                      </div>
                    </div>

                    <div className="card" style={{ marginBottom: '2rem' }}>
                      <h3 style={{ marginBottom: '1.5rem', display: 'flex', alignItems: 'center' }}>
                        <IndianRupee size={20} style={{ marginRight: '0.5rem', color: 'var(--accent)' }} /> 
                        Expense Breakdown
                      </h3>
                      <div style={{ display: 'flex', flexDirection: 'column', gap: '1rem' }}>
                        <div style={{ display: 'flex', justifyContent: 'space-between', paddingBottom: '0.5rem', borderBottom: '1px solid var(--border-color)' }}>
                          <span style={{ color: 'var(--text-secondary)' }}>Cleaning Costs</span>
                          <span style={{ fontWeight: '500' }}>₹{totalCleaningCosts.toFixed(2)}</span>
                        </div>
                        <div style={{ display: 'flex', justifyContent: 'space-between', paddingBottom: '0.5rem', borderBottom: '1px solid var(--border-color)' }}>
                          <span style={{ color: 'var(--text-secondary)' }}>Travel Expenses</span>
                          <span style={{ fontWeight: '500' }}>₹{totalTravelExpenses.toFixed(2)}</span>
                        </div>
                        <div style={{ display: 'flex', justifyContent: 'space-between', paddingBottom: '0.5rem', borderBottom: '1px solid var(--border-color)' }}>
                          <span style={{ color: 'var(--text-secondary)' }}>Other Expenses</span>
                          <span style={{ fontWeight: '500' }}>₹{totalOtherExpenses.toFixed(2)}</span>
                        </div>
                        <div style={{ display: 'flex', justifyContent: 'space-between', paddingTop: '0.5rem' }}>
                          <span style={{ fontWeight: '600' }}>Total Deductions</span>
                          <span style={{ fontWeight: '600', color: 'var(--danger)' }}>₹{(totalCleaningCosts + totalTravelExpenses + totalOtherExpenses).toFixed(2)}</span>
                        </div>
                      </div>
                    </div>

                    <div className="card" style={{ marginBottom: '2rem' }}>
                      <h3 style={{ marginBottom: '1.5rem', display: 'flex', alignItems: 'center' }}>
                        <Briefcase size={20} style={{ marginRight: '0.5rem', color: 'var(--accent)' }} /> 
                        Investment Management
                      </h3>
                      <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '2rem' }}>
                        <div className="form-group">
                          <label>{data.settings.owner1Name || 'ASHRAF M'} Investment (₹)</label>
                          <input
                            type="number"
                            value={data.settings.owner1Investment || 0}
                            onChange={e => updateSettings({ ...data.settings, owner1Investment: Number(e.target.value) })}
                          />
                        </div>
                        <div className="form-group">
                          <label>{data.settings.owner2Name || 'RASHEED M'} Investment (₹)</label>
                          <input
                            type="number"
                            value={data.settings.owner2Investment || 0}
                            onChange={e => updateSettings({ ...data.settings, owner2Investment: Number(e.target.value) })}
                          />
                        </div>
                      </div>
                    </div>

                    {/* Transaction Management Section */}
                    <div className="card" style={{ marginBottom: '2rem' }}>
                      <h3 style={{ marginBottom: '1.5rem', display: 'flex', alignItems: 'center' }}>
                        <Wallet size={20} style={{ marginRight: '0.5rem', color: 'var(--accent)' }} /> 
                        Owner Transactions (Investments, Withdrawals, Transfers)
                      </h3>
                      
                      <div className="card" style={{ background: 'var(--bg-color)', marginBottom: '1.5rem' }}>
                        <div className="dashboard-grid" style={{ gridTemplateColumns: 'repeat(auto-fit, minmax(140px, 1fr))', gap: '1rem' }}>
                          <div className="form-group" style={{ margin: 0 }}>
                            <label>From</label>
                            <select id="tx-from" className="form-group" style={{ width: '100%' }}>
                              <option value="BUSINESS">BUSINESS Pool</option>
                              <option value={data.settings.owner1Name || 'ASHRAF M'}>{data.settings.owner1Name || 'ASHRAF M'}</option>
                              <option value={data.settings.owner2Name || 'RASHEED M'}>{data.settings.owner2Name || 'RASHEED M'}</option>
                            </select>
                          </div>
                          <div className="form-group" style={{ margin: 0 }}>
                            <label>To</label>
                            <select id="tx-to" className="form-group" style={{ width: '100%' }}>
                              <option value={data.settings.owner1Name || 'ASHRAF M'}>{data.settings.owner1Name || 'ASHRAF M'}</option>
                              <option value={data.settings.owner2Name || 'RASHEED M'}>{data.settings.owner2Name || 'RASHEED M'}</option>
                              <option value="BUSINESS">BUSINESS Pool (Investment)</option>
                            </select>
                          </div>
                          <div className="form-group" style={{ margin: 0 }}>
                            <label>Amount (₹)</label>
                            <input type="number" id="tx-amount" placeholder="0.00" />
                          </div>
                          <div className="form-group" style={{ margin: 0 }}>
                            <label>Date</label>
                            <input type="date" id="tx-date" defaultValue={new Date().toISOString().split('T')[0]} />
                          </div>
                          <div className="form-group" style={{ margin: 0 }}>
                            <label>Payout Type</label>
                            <select id="tx-subtype" className="form-group" style={{ width: '100%' }}>
                              <option value="profit">Profit Payout</option>
                              <option value="investment">Investment Return</option>
                              <option value="transfer">Personal Transfer</option>
                            </select>
                          </div>
                          <div className="form-group" style={{ margin: 0, gridColumn: 'span 1' }}>
                            <label>Notes / Label</label>
                            <input type="text" id="tx-notes" placeholder="e.g., Personal Loan..." />
                          </div>
                          <div style={{ display: 'flex', alignItems: 'flex-end' }}>
                            <button className="btn-primary" style={{ width: '100%' }} onClick={() => {
                              const from = (document.getElementById('tx-from') as HTMLSelectElement).value
                              const to = (document.getElementById('tx-to') as HTMLSelectElement).value
                              const amount = Number((document.getElementById('tx-amount') as HTMLInputElement).value)
                              const date = (document.getElementById('tx-date') as HTMLInputElement).value
                              const notes = (document.getElementById('tx-notes') as HTMLInputElement).value
                              const subType = (document.getElementById('tx-subtype') as HTMLSelectElement).value as any
                              
                              if (amount > 0 && from !== to) {
                                handleAddOwnerTransaction({ from, to, amount, date, notes, subType })
                                ;(document.getElementById('tx-amount') as HTMLInputElement).value = ''
                                ;(document.getElementById('tx-notes') as HTMLInputElement).value = ''
                              } else if (from === to) {
                                alert("Source and destination must be different.")
                              }
                            }}>
                              Record Transaction
                            </button>
                          </div>
                        </div>
                      </div>

                      <div className="table-container">
                        <table className="data-table">
                          <thead>
                            <tr>
                              <th>Date</th>
                              <th>From</th>
                              <th>To</th>
                              <th>Type</th>
                              <th>Label/Notes</th>
                              <th style={{ textAlign: 'right' }}>Amount</th>
                              <th style={{ textAlign: 'center' }}>Action</th>
                            </tr>
                          </thead>
                          <tbody>
                            {(data.ownerTransactions || []).length === 0 ? (
                              <tr><td colSpan={6} style={{ textAlign: 'center', padding: '2rem', color: 'var(--text-secondary)' }}>No owner transactions found.</td></tr>
                            ) : (
                              [...data.ownerTransactions].sort((a, b) => new Date(b.date).getTime() - new Date(a.date).getTime()).map(t => (
                                <tr key={t.id}>
                                  <td>{t.date}</td>
                                  <td style={{ fontWeight: '500' }}>{t.from}</td>
                                  <td style={{ fontWeight: '500' }}>{t.to}</td>
                                  <td style={{ fontSize: '0.8rem', textTransform: 'capitalize' }}>{t.subType || '-'}</td>
                                  <td style={{ fontSize: '0.85rem' }}>{t.notes}</td>
                                  <td style={{ textAlign: 'right', fontWeight: '600', color: t.to === 'BUSINESS' ? 'var(--success)' : 'var(--danger)' }}>
                                    ₹{t.amount.toFixed(2)}
                                  </td>
                                  <td style={{ textAlign: 'center' }}>
                                    <button className="btn-secondary" style={{ padding: '0.3rem', color: 'var(--danger)' }} onClick={() => deleteOwnerTransaction(t.id)}>
                                      <Trash2 size={16} />
                                    </button>
                                  </td>
                                </tr>
                              ))
                            )}
                          </tbody>
                        </table>
                      </div>
                    </div>

                    <h2 style={{ marginBottom: '1.5rem' }}>Owner Balance Sheet (50/50 Split)</h2>
                    <div className="dashboard-grid">
                      {(() => {
                        const o1 = data.settings.owner1Name || 'ASHRAF M'
                        const o2 = data.settings.owner2Name || 'RASHEED M'
                        const txs = data.ownerTransactions || []

                        const getOwnerStats = (name: string, other: string) => {
                          const investments = txs.filter(t => t.from === name && t.to === 'BUSINESS').reduce((s, t) => s + t.amount, 0)
                          const legacyInvestment = name === o1 ? (data.settings.owner1Investment || 0) : (data.settings.owner2Investment || 0)
                          const totalInvested = investments + legacyInvestment
                          
                          const profitPayouts = txs.filter(t => t.from === 'BUSINESS' && t.to === name && t.subType === 'profit').reduce((s, t) => s + t.amount, 0)
                          const investmentReturns = txs.filter(t => t.from === 'BUSINESS' && t.to === name && t.subType === 'investment').reduce((s, t) => s + t.amount, 0)
                          const otherWithdrawals = txs.filter(t => t.from === 'BUSINESS' && t.to === name && !t.subType).reduce((s, t) => s + t.amount, 0)
                          
                          const transfersOut = txs.filter(t => t.from === name && t.to === other).reduce((s, t) => s + t.amount, 0)
                          const transfersIn = txs.filter(t => t.from === other && t.to === name).reduce((s, t) => s + t.amount, 0)
                          
                          const remainingInvestment = totalInvested - investmentReturns
                          const remainingProfit = ownerShare - profitPayouts - otherWithdrawals - transfersOut + transfersIn
                          
                          return { totalInvested, investmentReturns, remainingInvestment, ownerShare, profitPayouts, remainingProfit, totalStanding: remainingInvestment + remainingProfit }
                        }

                        const s1 = getOwnerStats(o1, o2)
                        const s2 = getOwnerStats(o2, o1)
                        
                        return (
                          <>
                            <div className="stat-card" style={{ borderTop: '4px solid #3b82f6' }}>
                              <div className="stat-title">{o1}</div>
                              <div style={{ marginBottom: '1rem' }}>
                                <div style={{ fontSize: '0.85rem', fontWeight: '600', color: 'var(--text-primary)', marginBottom: '0.5rem', borderBottom: '1px solid var(--border-color)' }}>Investment Tracking</div>
                                <div style={{ fontSize: '0.8rem', color: 'var(--text-secondary)', display: 'flex', justifyContent: 'space-between' }}>
                                  <span>Total Invested:</span> <span>₹{s1.totalInvested.toFixed(2)}</span>
                                </div>
                                <div style={{ fontSize: '0.8rem', color: 'var(--danger)', display: 'flex', justifyContent: 'space-between' }}>
                                  <span>Returned:</span> <span>-₹{s1.investmentReturns.toFixed(2)}</span>
                                </div>
                                <div style={{ fontSize: '0.85rem', color: 'var(--text-primary)', fontWeight: '500', display: 'flex', justifyContent: 'space-between', marginTop: '0.2rem' }}>
                                  <span>Net Investment:</span> <span>₹{s1.remainingInvestment.toFixed(2)}</span>
                                </div>

                                <div style={{ fontSize: '0.85rem', fontWeight: '600', color: 'var(--text-primary)', marginBottom: '0.5rem', marginTop: '1rem', borderBottom: '1px solid var(--border-color)' }}>Profit Tracking</div>
                                <div style={{ fontSize: '0.8rem', color: 'var(--text-secondary)', display: 'flex', justifyContent: 'space-between' }}>
                                  <span>Profit Share:</span> <span>₹{s1.ownerShare.toFixed(2)}</span>
                                </div>
                                <div style={{ fontSize: '0.8rem', color: 'var(--danger)', display: 'flex', justifyContent: 'space-between' }}>
                                  <span>Paid Out:</span> <span>-₹{s1.profitPayouts.toFixed(2)}</span>
                                </div>
                                <div style={{ fontSize: '0.85rem', color: 'var(--text-primary)', fontWeight: '500', display: 'flex', justifyContent: 'space-between', marginTop: '0.2rem' }}>
                                  <span>Remaining Profit:</span> <span>₹{s1.remainingProfit.toFixed(2)}</span>
                                </div>
                              </div>
                              <div className="stat-label" style={{ marginTop: '1rem' }}>Combined Balance</div>
                              <div className="stat-value" style={{ color: s1.totalStanding >= 0 ? 'var(--success)' : 'var(--danger)' }}>
                                ₹{s1.totalStanding.toFixed(2)}
                              </div>
                            </div>
                            
                            <div className="stat-card" style={{ borderTop: '4px solid #8b5cf6' }}>
                              <div className="stat-title">{o2}</div>
                              <div style={{ marginBottom: '1rem' }}>
                                <div style={{ fontSize: '0.85rem', fontWeight: '600', color: 'var(--text-primary)', marginBottom: '0.5rem', borderBottom: '1px solid var(--border-color)' }}>Investment Tracking</div>
                                <div style={{ fontSize: '0.8rem', color: 'var(--text-secondary)', display: 'flex', justifyContent: 'space-between' }}>
                                  <span>Total Invested:</span> <span>₹{s2.totalInvested.toFixed(2)}</span>
                                </div>
                                <div style={{ fontSize: '0.8rem', color: 'var(--danger)', display: 'flex', justifyContent: 'space-between' }}>
                                  <span>Returned:</span> <span>-₹{s2.investmentReturns.toFixed(2)}</span>
                                </div>
                                <div style={{ fontSize: '0.85rem', color: 'var(--text-primary)', fontWeight: '500', display: 'flex', justifyContent: 'space-between', marginTop: '0.2rem' }}>
                                  <span>Net Investment:</span> <span>₹{s2.remainingInvestment.toFixed(2)}</span>
                                </div>

                                <div style={{ fontSize: '0.85rem', fontWeight: '600', color: 'var(--text-primary)', marginBottom: '0.5rem', marginTop: '1rem', borderBottom: '1px solid var(--border-color)' }}>Profit Tracking</div>
                                <div style={{ fontSize: '0.8rem', color: 'var(--text-secondary)', display: 'flex', justifyContent: 'space-between' }}>
                                  <span>Profit Share:</span> <span>₹{s2.ownerShare.toFixed(2)}</span>
                                </div>
                                <div style={{ fontSize: '0.8rem', color: 'var(--danger)', display: 'flex', justifyContent: 'space-between' }}>
                                  <span>Paid Out:</span> <span>-₹{s2.profitPayouts.toFixed(2)}</span>
                                </div>
                                <div style={{ fontSize: '0.85rem', color: 'var(--text-primary)', fontWeight: '500', display: 'flex', justifyContent: 'space-between', marginTop: '0.2rem' }}>
                                  <span>Remaining Profit:</span> <span>₹{s2.remainingProfit.toFixed(2)}</span>
                                </div>
                              </div>
                              <div className="stat-label" style={{ marginTop: '1rem' }}>Combined Balance</div>
                              <div className="stat-value" style={{ color: s2.totalStanding >= 0 ? 'var(--success)' : 'var(--danger)' }}>
                                ₹{s2.totalStanding.toFixed(2)}
                              </div>
                            </div>
                          </>
                        )
                      })()}
                    </div>
                  </>
                )
              })()}
            </motion.div>
          )}

          {activeTab === 'about' && (
            <motion.div
              key="about"
              initial={{ opacity: 0, y: 20 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, y: -20 }}
            >
              <h1>About Melon Events</h1>
              <div className="card" style={{ maxWidth: '600px', textAlign: 'center', padding: '3rem 2rem' }}>
                <img src="logo-mark.png" alt="Melon Events" style={{ width: '80px', marginBottom: '1.5rem' }} />
                <h2 style={{ marginBottom: '0.5rem' }}>Melon Events</h2>
                <div style={{ color: 'var(--text-secondary)', marginBottom: '2rem' }}>Version 1.0.4 (Stable)</div>

                <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '2rem', textAlign: 'left', marginBottom: '3rem' }}>
                  <div className="card" style={{ padding: '1.5rem', background: 'var(--bg-color)', margin: 0 }}>
                    <div style={{ fontSize: '0.8rem', color: 'var(--text-secondary)', textTransform: 'uppercase', marginBottom: '0.5rem' }}>Developer</div>
                    <div style={{ fontWeight: '700', fontSize: '1.1rem' }}>Alinshan</div>
                  </div>
                  <div className="card" style={{ padding: '1.5rem', background: 'var(--bg-color)', margin: 0 }}>
                    <div style={{ fontSize: '0.8rem', color: 'var(--text-secondary)', textTransform: 'uppercase', marginBottom: '0.5rem' }}>License</div>
                    <div style={{ fontWeight: '700', fontSize: '1.1rem' }}>MIT License</div>
                    <button
                      className="btn-secondary"
                      style={{ padding: '0.2rem 0.5rem', fontSize: '0.7rem', marginTop: '0.5rem', width: 'auto' }}
                      onClick={() => setShowLicense(true)}
                    >
                      View Text
                    </button>
                  </div>
                </div>

                <div style={{ borderTop: '1px solid var(--border-color)', paddingTop: '2rem' }}>
                  <h3 style={{ marginBottom: '1rem' }}>Updates & Support</h3>
                  <p style={{ color: 'var(--text-secondary)', marginBottom: '1.5rem', fontSize: '0.9rem' }}>
                    Click the button below to check for the latest versions and documentation on the official GitHub repository.
                  </p>
                  <div style={{ display: 'flex', gap: '1rem', justifyContent: 'center' }}>
                    <button className="btn-primary" onClick={() => (window as any).electronAPI.openExternal('https://github.com/Alinshan/Melon-Events/releases')}>
                      <ExternalLink size={18} /> Visit GitHub Releases
                    </button>
                    <button className="btn-secondary" onClick={async () => {
                      const res = await (window as any).electronAPI.checkForUpdates()
                      if (res.status === 'error') alert('Update check failed: ' + res.message)
                      else if (res.status === 'dev') alert(res.message)
                    }}>
                      Check for Updates
                    </button>
                  </div>
                </div>

                <div style={{ marginTop: '3rem', fontSize: '0.8rem', color: 'var(--text-secondary)' }}>
                  © 2026 Melon Events | Developed by Alinshan
                </div>
              </div>

              {/* Full License Modal */}
              {showLicense && (
                <div style={{
                  position: 'fixed', top: 0, left: 0, right: 0, bottom: 0,
                  background: 'rgba(15, 23, 42, 0.9)', backdropFilter: 'blur(8px)',
                  display: 'flex', alignItems: 'center', justifyContent: 'center', zIndex: 2000,
                  padding: '2rem'
                }}>
                  <div className="card" style={{ maxWidth: '600px', width: '100%', maxHeight: '80vh', overflowY: 'auto' }}>
                    <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '1.5rem' }}>
                      <h3>MIT License</h3>
                      <button className="btn-secondary" onClick={() => setShowLicense(false)}>Close</button>
                    </div>
                    <pre style={{
                      whiteSpace: 'pre-wrap',
                      fontSize: '0.85rem',
                      lineHeight: '1.6',
                      color: 'var(--text-primary)',
                      fontFamily: 'monospace',
                      textAlign: 'left',
                      background: 'rgba(0,0,0,0.2)',
                      padding: '1.5rem',
                      borderRadius: '12px'
                    }}>
                      {`MIT License

Copyright (c) 2026 Alinshan

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.`}
                    </pre>
                  </div>
                </div>
              )}
            </motion.div>
          )}
        </AnimatePresence>
      </main>

      {/* Return Modal Overlay */}
      {returningRentalId && (
        <div style={{
          position: 'fixed', top: 0, left: 0, right: 0, bottom: 0,
          background: 'rgba(15, 23, 42, 0.8)', backdropFilter: 'blur(4px)',
          display: 'flex', alignItems: 'center', justifyContent: 'center', zIndex: 1000
        }}>
          <form className="card" style={{ width: '400px', border: '1px solid var(--border-color)' }} onSubmit={confirmReturn}>
            <h3>Confirm Return</h3>
            <p style={{ color: 'var(--text-secondary)', marginBottom: '1.5rem', fontSize: '0.9rem' }}>
              Are there any damage or missing item charges for this rental?
            </p>
            <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '1rem', marginBottom: '1.5rem' }}>
              <div className="form-group">
                <label>Unused Plates</label>
                <input
                  type="number"
                  min="0"
                  value={unusedPlatesInput}
                  onChange={e => setUnusedPlatesInput(e.target.value)}
                />
              </div>
              <div className="form-group">
                <label>Unused Glasses</label>
                <input
                  type="number"
                  min="0"
                  value={unusedGlassesInput}
                  onChange={e => setUnusedGlassesInput(e.target.value)}
                />
              </div>
              <div className="form-group" style={{ gridColumn: 'span 2' }}>
                <label>Damage Charges (₹)</label>
                <input
                  type="number"
                  min="0"
                  step="0.01"
                  value={damageChargeInput}
                  onChange={e => setDamageChargeInput(e.target.value)}
                  autoFocus
                />
              </div>
            </div>
            <div style={{ display: 'flex', gap: '1rem', marginTop: '1.5rem', justifyContent: 'flex-end' }}>
              <button type="button" className="btn-secondary" onClick={cancelReturn}>Cancel</button>
              <button type="submit" className="btn-primary">Confirm Return</button>
            </div>
          </form>
        </div>
      )}

      {/* Cleaning Modal Overlay */}
      {cleaningRentalId && (
        <div style={{
          position: 'fixed', top: 0, left: 0, right: 0, bottom: 0,
          background: 'rgba(15, 23, 42, 0.8)', backdropFilter: 'blur(4px)',
          display: 'flex', alignItems: 'center', justifyContent: 'center', zIndex: 1000
        }}>
          <form className="card" style={{ width: '400px', border: '1px solid var(--border-color)' }} onSubmit={confirmCleaning}>
            <h3>Assign Cleaning Staff</h3>
            <p style={{ color: 'var(--text-secondary)', marginBottom: '1.5rem', fontSize: '0.9rem' }}>
              Select an employee to assign this cleaning task.
            </p>
            <div className="form-group" style={{ marginBottom: '1.5rem' }}>
              <label>Employee Name</label>
              <select
                value={cleaningEmployeeSelect}
                onChange={e => setCleaningEmployeeSelect(e.target.value)}
                style={{ width: '100%', padding: '0.75rem', borderRadius: '4px', background: 'var(--bg-color)', border: '1px solid var(--border-color)', color: 'var(--text-color)' }}
              >
                <option value="">-- Select Employee --</option>
                {data.settings.employees?.map(emp => (
                  <option key={emp} value={emp}>{emp}</option>
                ))}
              </select>
            </div>

            <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '1rem', marginBottom: '1.5rem' }}>
              <div className="form-group">
                <label>Plates to Clean</label>
                <input
                  type="number"
                  min="0"
                  value={cleaningPlatesInput}
                  onChange={e => setCleaningPlatesInput(e.target.value)}
                />
              </div>
              <div className="form-group">
                <label>Glasses to Clean</label>
                <input
                  type="number"
                  min="0"
                  value={cleaningGlassesInput}
                  onChange={e => setCleaningGlassesInput(e.target.value)}
                />
              </div>
              <div className="form-group">
                <label>Plate Rate (₹)</label>
                <input
                  type="number"
                  min="0"
                  step="0.01"
                  value={cleaningPlateRateInput}
                  onChange={e => setCleaningPlateRateInput(e.target.value)}
                />
              </div>
              <div className="form-group">
                <label>Glass Rate (₹)</label>
                <input
                  type="number"
                  min="0"
                  step="0.01"
                  value={cleaningGlassRateInput}
                  onChange={e => setCleaningGlassRateInput(e.target.value)}
                />
              </div>
            </div>

            <div style={{ fontSize: '1.1rem', fontWeight: '700', color: 'var(--success)', textAlign: 'right', marginBottom: '1rem' }}>
              Total Salary: ₹{((Number(cleaningPlatesInput) * Number(cleaningPlateRateInput)) + (Number(cleaningGlassesInput) * Number(cleaningGlassRateInput))).toFixed(2)}
            </div>
            <div style={{ display: 'flex', gap: '1rem', marginTop: '1.5rem', justifyContent: 'flex-end' }}>
              <button type="button" className="btn-secondary" onClick={cancelCleaning}>Cancel</button>
              <button type="submit" className="btn-primary">Confirm Cleaning</button>
            </div>
          </form>
        </div>
      )}

      {editingRental && (
        <EditRentalModal
          rental={editingRental}
          settings={data.settings}
          onSave={updateRental}
          onCancel={() => setEditingRental(null)}
        />
      )}

      {editingSalary && (
        <EditSalaryModal
          salary={editingSalary}
          onSave={updateSalary}
          onCancel={() => setEditingSalary(null)}
        />
      )}

      {editingExpense && (
        <EditExpenseModal
          expense={editingExpense}
          onSave={updateExpense}
          onCancel={() => setEditingExpense(null)}
        />
      )}
    </div>
  )
}

// --- Sub-components ---

function NewRentalForm({ settings, onSubmit }: { settings: ItemSettings, onSubmit: any }) {
  const [customer, setCustomer] = useState('')
  const [plates, setPlates] = useState<string | number>('')
  const [glasses, setGlasses] = useState<string | number>('')
  const [platePriceOverride, setPlatePriceOverride] = useState<string | number>(settings.platePrice)
  const [glassPriceOverride, setGlassPriceOverride] = useState<string | number>(settings.glassPrice)
  const [discount, setDiscount] = useState<string | number>('')
  const [travelExpense, setTravelExpense] = useState<string | number>('')
  const [date, setDate] = useState(new Date().toISOString().split('T')[0])

  const subtotal = (Number(plates) * Number(platePriceOverride)) + (Number(glasses) * Number(glassPriceOverride))
  const total = subtotal - Number(discount)

  const handleSubmit = (e: React.FormEvent) => {
    e.preventDefault()
    if (!customer || (plates === 0 && glasses === 0)) return
    onSubmit({
      customerName: customer,
      plateCount: plates,
      glassCount: glasses,
      platePrice: Number(platePriceOverride),
      glassPrice: Number(glassPriceOverride),
      discount,
      travelExpense,
      total,
      date
    })
    setCustomer('')
    setPlates('')
    setGlasses('')
    setDiscount('')
    setTravelExpense('')
  }

  return (
    <form className="card" onSubmit={handleSubmit}>
      <h3>Create New Invoice</h3>
      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(200px, 1fr))', gap: '1rem', marginTop: '1rem' }}>
        <div className="form-group">
          <label>Date</label>
          <input type="date" value={date} onChange={e => setDate(e.target.value)} />
        </div>
        <div className="form-group">
          <label>Customer Name</label>
          <input type="text" value={customer} onChange={e => setCustomer(e.target.value)} placeholder="e.g. Wedding Event" />
        </div>
        <div className="form-group">
          <label>Plate Count</label>
          <input type="number" min="0" value={plates} onChange={e => setPlates(e.target.value)} />
        </div>
        <div className="form-group">
          <label>Plate Price (₹)</label>
          <input type="number" min="0" step="0.01" value={platePriceOverride} onChange={e => setPlatePriceOverride(e.target.value)} />
        </div>
        <div className="form-group">
          <label>Glass Count</label>
          <input type="number" min="0" value={glasses} onChange={e => setGlasses(e.target.value)} />
        </div>
        <div className="form-group">
          <label>Glass Price (₹)</label>
          <input type="number" min="0" step="0.01" value={glassPriceOverride} onChange={e => setGlassPriceOverride(e.target.value)} />
        </div>
        <div className="form-group">
          <label>Discount (₹)</label>
          <input type="number" min="0" step="0.01" value={discount} onChange={e => setDiscount(e.target.value)} />
        </div>
        <div className="form-group">
          <label>Travel Expense (₹)</label>
          <input type="number" min="0" step="0.01" value={travelExpense} onChange={e => setTravelExpense(e.target.value)} placeholder="Internal cost" />
        </div>
      </div>
      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginTop: '1rem' }}>
        <div style={{ fontSize: '1.25rem' }}>
          Customer Total: <span style={{ color: 'var(--success)', fontWeight: '700' }}>₹{total.toFixed(2)}</span>
        </div>
        <button type="submit" className="btn-primary">
          <Plus size={18} /> Create Rental
        </button>
      </div>
    </form>
  )
}


// --- Booking Calendar Component ---
function BookingCalendar({ rentals }: { rentals: Rental[] }) {
  const today = new Date()
  const [viewDate, setViewDate] = useState(new Date(today.getFullYear(), today.getMonth(), 1))
  const [selectedDay, setSelectedDay] = useState<string | null>(null)

  const year = viewDate.getFullYear()
  const month = viewDate.getMonth()

  // Build a map of date -> rentals
  const rentalMap = useMemo(() => {
    const map: Record<string, Rental[]> = {}
    rentals.forEach(r => {
      if (!map[r.date]) map[r.date] = []
      map[r.date].push(r)
    })
    return map
  }, [rentals])

  const monthName = viewDate.toLocaleString('default', { month: 'long', year: 'numeric' })
  const firstDay = new Date(year, month, 1).getDay() // 0=Sun
  const daysInMonth = new Date(year, month + 1, 0).getDate()

  const prevMonth = () => setViewDate(new Date(year, month - 1, 1))
  const nextMonth = () => setViewDate(new Date(year, month + 1, 1))

  const selectedRentals = selectedDay ? (rentalMap[selectedDay] || []) : []

  const dayNames = ['Su', 'Mo', 'Tu', 'We', 'Th', 'Fr', 'Sa']

  return (
    <div className="card" style={{ padding: '1.25rem', position: 'sticky', top: '0' }}>
      {/* Calendar Header */}
      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '1rem' }}>
        <button
          className="btn-secondary"
          style={{ padding: '0.4rem 0.75rem', fontSize: '1rem' }}
          onClick={prevMonth}
        >‹</button>
        <span style={{ fontWeight: '700', fontSize: '0.95rem' }}>{monthName}</span>
        <button
          className="btn-secondary"
          style={{ padding: '0.4rem 0.75rem', fontSize: '1rem' }}
          onClick={nextMonth}
        >›</button>
      </div>

      {/* Day Headers */}
      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(7, 1fr)', gap: '2px', marginBottom: '4px' }}>
        {dayNames.map(d => (
          <div key={d} style={{
            textAlign: 'center',
            fontSize: '0.7rem',
            fontWeight: '600',
            color: 'var(--text-secondary)',
            padding: '4px 0'
          }}>{d}</div>
        ))}
      </div>

      {/* Calendar Grid */}
      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(7, 1fr)', gap: '3px' }}>
        {/* Empty cells for first day offset */}
        {Array.from({ length: firstDay }).map((_, i) => (
          <div key={`empty-${i}`} />
        ))}

        {/* Day cells */}
        {Array.from({ length: daysInMonth }).map((_, i) => {
          const day = i + 1
          const dateStr = `${year}-${String(month + 1).padStart(2, '0')}-${String(day).padStart(2, '0')}`
          const dayRentals = rentalMap[dateStr] || []
          const hasActive = dayRentals.some(r => !r.isReturned)
          const hasReturned = dayRentals.some(r => r.isReturned)
          const isToday = dateStr === today.toISOString().split('T')[0]
          const isSelected = selectedDay === dateStr

          let dotColor = ''
          if (hasActive) dotColor = 'var(--warning)'
          else if (hasReturned) dotColor = 'var(--success)'

          return (
            <div
              key={dateStr}
              onClick={() => setSelectedDay(isSelected ? null : dateStr)}
              style={{
                textAlign: 'center',
                padding: '5px 2px',
                borderRadius: '8px',
                cursor: dayRentals.length > 0 ? 'pointer' : 'default',
                fontSize: '0.8rem',
                fontWeight: isToday ? '800' : '400',
                color: isToday ? 'var(--accent-color)' : isSelected ? '#ffffff' : 'var(--text-primary)',
                background: isSelected
                  ? 'var(--accent-color)'
                  : isToday
                    ? 'rgba(56, 189, 248, 0.12)'
                    : dayRentals.length > 0
                      ? 'var(--border-color)'
                      : 'transparent',
                border: isToday && !isSelected ? '1px solid var(--accent-color)' : '1px solid transparent',
                transition: 'all 0.15s ease',
                position: 'relative',
              }}
            >
              {day}
              {dotColor && (
                <div style={{
                  width: '5px', height: '5px',
                  borderRadius: '50%',
                  background: isSelected ? '#ffffff' : dotColor,
                  margin: '2px auto 0'
                }} />
              )}
            </div>
          )
        })}
      </div>

      {/* Legend */}
      <div style={{ display: 'flex', gap: '1rem', marginTop: '1rem', justifyContent: 'center' }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: '5px', fontSize: '0.7rem', color: 'var(--text-secondary)' }}>
          <div style={{ width: '8px', height: '8px', borderRadius: '50%', background: 'var(--warning)' }} />
          Active
        </div>
        <div style={{ display: 'flex', alignItems: 'center', gap: '5px', fontSize: '0.7rem', color: 'var(--text-secondary)' }}>
          <div style={{ width: '8px', height: '8px', borderRadius: '50%', background: 'var(--success)' }} />
          Returned
        </div>
      </div>

      {/* Selected Day Bookings */}
      {selectedDay && (
        <div style={{ marginTop: '1rem', borderTop: '1px solid var(--border-color)', paddingTop: '1rem' }}>
          <div style={{ fontSize: '0.8rem', fontWeight: '700', color: 'var(--text-secondary)', marginBottom: '0.5rem' }}>
            {new Date(selectedDay + 'T00:00:00').toLocaleDateString('en-IN', { day: 'numeric', month: 'short', year: 'numeric' })}
          </div>
          {selectedRentals.length === 0 ? (
            <div style={{ fontSize: '0.8rem', color: 'var(--text-secondary)' }}>No bookings on this day.</div>
          ) : (
            selectedRentals.map(r => (
              <div key={r.id} style={{
                background: 'var(--bg-color)',
                border: '1px solid var(--border-color)',
                borderRadius: '10px',
                padding: '0.6rem 0.75rem',
                marginBottom: '0.4rem',
                borderLeft: `3px solid ${r.isReturned ? 'var(--success)' : 'var(--warning)'}`
              }}>
                <div style={{ fontWeight: '600', fontSize: '0.82rem' }}>{r.customerName}</div>
                <div style={{ fontSize: '0.75rem', color: 'var(--text-secondary)' }}>
                  {r.plateCount}P / {r.glassCount}G &nbsp;·&nbsp;
                  <span style={{ color: 'var(--text-primary)', fontWeight: '700' }}>₹{r.total}</span>
                </div>
              </div>
            ))
          )}
        </div>
      )}
    </div>
  )
}

function ExpenseForm({ onSubmit }: { onSubmit: (record: Omit<ExpenseRecord, 'id'>) => void }) {
  const [date, setDate] = useState(new Date().toISOString().split('T')[0])
  const [category, setCategory] = useState('Travel & Transport')
  const [description, setDescription] = useState('')
  const [amount, setAmount] = useState<number | ''>('')

  const categories = ['Travel & Transport', 'Packaging', 'Cleaning Solutions', 'Maintenance', 'Damage/Loss', 'Other']

  const handleSubmit = (e: React.FormEvent) => {
    e.preventDefault()
    if (!amount || Number(amount) <= 0) return
    onSubmit({
      date,
      category,
      description,
      amount: Number(amount)
    })
    setDescription('')
    setAmount('')
  }

  return (
    <form className="card" onSubmit={handleSubmit} style={{ marginBottom: '1.5rem' }}>
      <h3>Log Business Expense</h3>
      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(200px, 1fr))', gap: '1rem', marginTop: '1rem' }}>
        <div className="form-group">
          <label>Date</label>
          <input type="date" value={date} onChange={e => setDate(e.target.value)} required />
        </div>
        <div className="form-group">
          <label>Category</label>
          <select value={category} onChange={e => setCategory(e.target.value)} required>
            {categories.map(c => <option key={c} value={c}>{c}</option>)}
          </select>
        </div>
        <div className="form-group">
          <label>Description</label>
          <input type="text" value={description} onChange={e => setDescription(e.target.value)} placeholder="E.g., Bought soap, Auto fare" required />
        </div>
        <div className="form-group">
          <label>Amount (₹)</label>
          <input type="number" step="0.01" min="0.01" value={amount} onChange={e => setAmount(e.target.value === '' ? '' : Number(e.target.value))} placeholder="0.00" required />
        </div>
      </div>
      <div style={{ display: 'flex', justifyContent: 'flex-end', marginTop: '1rem' }}>
        <button type="submit" className="btn-primary" style={{ background: 'var(--danger)' }}>
          <Wallet size={18} /> Record Expense
        </button>
      </div>
    </form>
  )
}

export default App
