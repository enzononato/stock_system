import { Routes, Route, Navigate } from 'react-router-dom'
import { useAuth } from '@/contexts/AuthContext'
import AppLayout from '@/components/layout/AppLayout'
import { RestrictedAccess } from '@/components/layout/RestrictedAccess'
import LoginPage from '@/pages/LoginPage'
import EsqueciSenhaPage from '@/pages/EsqueciSenhaPage'
import RedefinirSenhaPage from '@/pages/RedefinirSenhaPage'
import PerfilPage from '@/pages/PerfilPage'
import DashboardPage from '@/pages/DashboardPage'
import StockPage from '@/pages/StockPage'
import ItemDetailPage from '@/pages/ItemDetailPage'
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
import UnidadesPage from '@/pages/UnidadesPage'
import ColaboradoresPage from '@/pages/ColaboradoresPage'
import ColaboradorDetailPage from '@/pages/ColaboradorDetailPage'
import AuditPage from '@/pages/AuditPage'
import LocaisPage from '@/pages/LocaisPage'
import LinhasPage from '@/pages/LinhasPage'
import LinhasRelatorioPage from '@/pages/LinhasRelatorioPage'
import EtiquetasPage from '@/pages/EtiquetasPage'
import EtiquetaRedirectPage from '@/pages/EtiquetaRedirectPage'
import InventariosPage from '@/pages/InventariosPage'
import InventarioLeituraPage from '@/pages/InventarioLeituraPage'
import InventarioRelatorioPage from '@/pages/InventarioRelatorioPage'
import { PosLoginRedirect } from '@/features/etiquetas/PosLoginRedirect'
import SistemasPage from '@/pages/acessos/SistemasPage'
import PacotesPage from '@/pages/acessos/PacotesPage'
import AcessoEventosPage from '@/pages/acessos/AcessoEventosPage'
import AcessoEventoDetailPage from '@/pages/acessos/AcessoEventoDetailPage'
import ManutencoesPage from '@/pages/ManutencoesPage'
import { ToastContainer } from '@/components/ui/toast'

/**
 * Guard de papel. Antes redirecionava silenciosamente para "/" quando o
 * usuário não tinha o papel necessário — o usuário clicava em algo, caía no
 * Dashboard e não entendia por quê. Agora mostra `RestrictedAccess`
 * explicando o motivo e oferecendo um caminho de volta (Task 8, Step 4).
 * Sem usuário logado ainda vale a pena mandar para /login — não há nada a
 * explicar nesse caso, é só sessão ausente.
 */
function RequireRole({ roles, children }: { roles: string[]; children: React.ReactNode }) {
  const { user } = useAuth()
  if (!user) return <Navigate to="/login" replace />
  if (!roles.includes(user.role)) return <RestrictedAccess roles={roles} currentRole={user.role} />
  return <>{children}</>
}

export default function App() {
  return (
    <>
      <Routes>
        <Route path="/login" element={<LoginPage />} />
        <Route path="/e/:id" element={<EtiquetaRedirectPage />} />
        <Route path="/esqueci-senha" element={<EsqueciSenhaPage />} />
        <Route path="/redefinir-senha" element={<RedefinirSenhaPage />} />

        {/* Todas as rotas protegidas dentro do AppLayout */}
        <Route element={<AppLayout />}>
          {/* Acessível por todos os roles — Dashboard é a nova home (Task 7);
              o estoque, que ocupava "/", mudou para "/stock". */}
          <Route index element={<DashboardPage />} />
          <Route path="stock" element={<StockPage />} />
          <Route path="items/:id" element={<ItemDetailPage />} />
          <Route path="charts" element={<ChartsPage />} />
          <Route path="perfil" element={<PerfilPage />} />

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

          <Route
            path="colaboradores"
            element={
              <RequireRole roles={['Gestor', 'Técnico']}>
                <ColaboradoresPage />
              </RequireRole>
            }
          />
          <Route
            path="colaboradores/:id"
            element={
              <RequireRole roles={['Gestor', 'Técnico']}>
                <ColaboradorDetailPage />
              </RequireRole>
            }
          />

          <Route
            path="locais"
            element={
              <RequireRole roles={['Gestor', 'Técnico']}>
                <LocaisPage />
              </RequireRole>
            }
          />
          <Route
            path="etiquetas"
            element={
              <RequireRole roles={['Gestor', 'Técnico']}>
                <EtiquetasPage />
              </RequireRole>
            }
          />
          <Route
            path="inventarios"
            element={
              <RequireRole roles={['Gestor', 'Técnico']}>
                <InventariosPage />
              </RequireRole>
            }
          />
          <Route
            path="inventarios/:id/leitura"
            element={
              <RequireRole roles={['Gestor', 'Técnico']}>
                <InventarioLeituraPage />
              </RequireRole>
            }
          />
          <Route
            path="inventarios/:id"
            element={
              <RequireRole roles={['Gestor', 'Técnico']}>
                <InventarioRelatorioPage />
              </RequireRole>
            }
          />

          <Route path="linhas" element={<LinhasPage />} />
          <Route
            path="linhas/relatorio"
            element={
              <RequireRole roles={['Gestor', 'Técnico']}>
                <LinhasRelatorioPage />
              </RequireRole>
            }
          />

          <Route
            path="manutencoes"
            element={
              <RequireRole roles={['Gestor', 'Técnico']}>
                <ManutencoesPage />
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
          <Route
            path="unidades"
            element={
              <RequireRole roles={['Gestor']}>
                <UnidadesPage />
              </RequireRole>
            }
          />

          <Route
            path="auditoria"
            element={
              <RequireRole roles={['Gestor']}>
                <AuditPage />
              </RequireRole>
            }
          />

          <Route
            path="acessos/sistemas"
            element={
              <RequireRole roles={['Gestor']}>
                <SistemasPage />
              </RequireRole>
            }
          />

          <Route
            path="acessos/pacotes"
            element={
              <RequireRole roles={['Gestor']}>
                <PacotesPage />
              </RequireRole>
            }
          />

          <Route
            path="acessos/eventos"
            element={
              <RequireRole roles={['Gestor', 'Técnico']}>
                <AcessoEventosPage />
              </RequireRole>
            }
          />
          <Route path="acessos" element={<Navigate to="/acessos/eventos" replace />} />

          <Route
            path="acessos/eventos/:id"
            element={
              <RequireRole roles={['Gestor', 'Técnico']}>
                <AcessoEventoDetailPage />
              </RequireRole>
            }
          />

          {/* Fallback */}
          <Route path="*" element={<Navigate to="/" replace />} />
        </Route>
      </Routes>

      <PosLoginRedirect />
      <ToastContainer />
    </>
  )
}
