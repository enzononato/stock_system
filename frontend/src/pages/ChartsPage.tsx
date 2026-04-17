import { useState } from 'react'
import { useQuery } from '@tanstack/react-query'
import { getLoansChart, getRegistrationsChart } from '@/api/reports'
import { BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer } from 'recharts'
import { Button } from '@/components/ui/button'
import { Input } from '@/components/ui/input'
import { Label } from '@/components/ui/label'
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from '@/components/ui/select'

const MONTHS = ['Janeiro','Fevereiro','Março','Abril','Maio','Junho','Julho','Agosto','Setembro','Outubro','Novembro','Dezembro']

export default function ChartsPage() {
  const now = new Date()
  const [year, setYear] = useState(String(now.getFullYear()))
  const [month, setMonth] = useState(String(now.getMonth() + 1))
  const [params, setParams] = useState({ year: now.getFullYear(), month: now.getMonth() + 1 })

  const { data: loansData } = useQuery({
    queryKey: ['chart-loans', params.year, params.month],
    queryFn: () => getLoansChart(params.year, params.month),
  })

  const { data: registrationsData } = useQuery({
    queryKey: ['chart-registrations', params.year, params.month],
    queryFn: () => getRegistrationsChart(params.year, params.month),
  })

  function buildLoansChartData() {
    if (!loansData) return []
    return loansData.days.map((day, i) => ({
      dia: `Dia ${day}`,
      Empréstimos: loansData.values[i],
      Devoluções: loansData.values2?.[i] ?? 0,
    }))
  }

  function buildRegChartData() {
    if (!registrationsData) return []
    return registrationsData.days.map((day, i) => ({
      dia: `Dia ${day}`,
      Cadastros: registrationsData.values[i],
    }))
  }

  return (
    <div className="space-y-8">
      <div>
        <h2 className="text-xl font-semibold">Gráficos</h2>
        <p className="text-sm text-slate-500">Visualize as tendências mensais de empréstimos e cadastros.</p>
      </div>

      <div className="flex items-end gap-4 bg-white rounded-xl border p-4">
        <div className="flex flex-col gap-1.5">
          <Label>Ano</Label>
          <Input value={year} onChange={e => setYear(e.target.value)} className="w-24" />
        </div>
        <div className="flex flex-col gap-1.5">
          <Label>Mês</Label>
          <Select value={month} onValueChange={setMonth}>
            <SelectTrigger className="w-40"><SelectValue /></SelectTrigger>
            <SelectContent>
              {MONTHS.map((m, i) => <SelectItem key={i+1} value={String(i+1)}>{m}</SelectItem>)}
            </SelectContent>
          </Select>
        </div>
        <Button onClick={() => setParams({ year: Number(year), month: Number(month) })}>
          Atualizar
        </Button>
      </div>

      {/* Gráfico 1: Empréstimos x Devoluções */}
      <div className="bg-white rounded-xl border p-6">
        <h3 className="font-medium text-slate-700 mb-4">Empréstimos × Devoluções por Dia</h3>
        <ResponsiveContainer width="100%" height={300}>
          <BarChart data={buildLoansChartData()} margin={{ top: 5, right: 20, left: 0, bottom: 5 }}>
            <CartesianGrid strokeDasharray="3 3" stroke="#f0f0f0" />
            <XAxis dataKey="dia" tick={{ fontSize: 11 }} />
            <YAxis allowDecimals={false} />
            <Tooltip />
            <Legend />
            <Bar dataKey="Empréstimos" fill="#3b82f6" radius={[4, 4, 0, 0]} />
            <Bar dataKey="Devoluções" fill="#10b981" radius={[4, 4, 0, 0]} />
          </BarChart>
        </ResponsiveContainer>
      </div>

      {/* Gráfico 2: Cadastros */}
      <div className="bg-white rounded-xl border p-6">
        <h3 className="font-medium text-slate-700 mb-4">Novos Cadastros por Dia</h3>
        <ResponsiveContainer width="100%" height={300}>
          <BarChart data={buildRegChartData()} margin={{ top: 5, right: 20, left: 0, bottom: 5 }}>
            <CartesianGrid strokeDasharray="3 3" stroke="#f0f0f0" />
            <XAxis dataKey="dia" tick={{ fontSize: 11 }} />
            <YAxis allowDecimals={false} />
            <Tooltip />
            <Bar dataKey="Cadastros" fill="#8b5cf6" radius={[4, 4, 0, 0]} />
          </BarChart>
        </ResponsiveContainer>
      </div>
    </div>
  )
}
