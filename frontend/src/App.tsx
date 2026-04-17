import { Routes, Route, Navigate } from 'react-router-dom'
import { useAuth } from '@/contexts/AuthContext'
import AppLayout from '@/components/layout/AppLayout'
import LoginPage from '@/pages/LoginPage'
import StockPage from '@/pages/StockPage'
import RegisterItemPage from '@/pages/RegisterItemPage'
import PeripheralsPage from '@/pages/PeripheralsPage'
import LinkPeripheralPage from '@/pages/LinkPeripheralPage'
import LoanPage from '@/pages/LoanPage'
import ReturnPage from '@/pages/ReturnPage'
import RemovePage from '@/pages/RemovePage'
import HistoryPage from '@/pages/HistoryPage'
import ReportPage from '@/pages/ReportPage'
import ChartsPage from '@/pages/ChartsPage'
import TermsPage from '@/pages/TermsPage'
import UsersPage from '@/pages/UsersPage'
import { ToastContainer } from '@/components/ui/toast'

/** Guard que redireciona para /login se não tiver o role necessário */
function RequireRole({ roles, children }: { roles: string[]; children: React.ReactNode }) {
  const { user } = useAuth()
  if (!user || !roles.includes(user.role)) return <Navigate to="/" replace />
  return <>{children}</>
}

export default function App() {
  return (
    <>
      <Routes>
        <Route path="/login" element={<LoginPage />} />

        {/* Todas as rotas protegidas dentro do AppLayout */}
        <Route element={<AppLayout />}>
          {/* Acessível por todos os roles */}
          <Route index element={<StockPage />} />
          <Route path="charts" element={<ChartsPage />} />

          {/* Gestor + Técnico */}
          <Route
            path="register"
            element={
              <RequireRole roles={['Gestor', 'Técnico']}>
                <RegisterItemPage mode="create" />
              </RequireRole>
            }
          />
          <Route
            path="edit/:id"
            element={
              <RequireRole roles={['Gestor', 'Técnico']}>
                <RegisterItemPage mode="edit" />
              </RequireRole>
            }
          />
          <Route
            path="peripherals"
            element={
              <RequireRole roles={['Gestor', 'Técnico']}>
                <PeripheralsPage />
              </RequireRole>
            }
          />
          <Route
            path="link"
            element={
              <RequireRole roles={['Gestor', 'Técnico']}>
                <LinkPeripheralPage />
              </RequireRole>
            }
          />
          <Route
            path="loan"
            element={
              <RequireRole roles={['Gestor', 'Técnico']}>
                <LoanPage />
              </RequireRole>
            }
          />
          <Route
            path="return"
            element={
              <RequireRole roles={['Gestor', 'Técnico']}>
                <ReturnPage />
              </RequireRole>
            }
          />
          <Route
            path="history"
            element={
              <RequireRole roles={['Gestor', 'Técnico']}>
                <HistoryPage />
              </RequireRole>
            }
          />
          <Route
            path="report"
            element={
              <RequireRole roles={['Gestor', 'Técnico']}>
                <ReportPage />
              </RequireRole>
            }
          />
          <Route
            path="terms"
            element={
              <RequireRole roles={['Gestor', 'Técnico']}>
                <TermsPage />
              </RequireRole>
            }
          />

          {/* Somente Gestor */}
          <Route
            path="remove"
            element={
              <RequireRole roles={['Gestor']}>
                <RemovePage />
              </RequireRole>
            }
          />
          <Route
            path="users"
            element={
              <RequireRole roles={['Gestor']}>
                <UsersPage />
              </RequireRole>
            }
          />

          {/* Fallback */}
          <Route path="*" element={<Navigate to="/" replace />} />
        </Route>
      </Routes>

      <ToastContainer />
    </>
  )
}
