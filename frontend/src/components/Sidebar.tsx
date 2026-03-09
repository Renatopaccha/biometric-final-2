import { LayoutDashboard, Database, BarChart3, FlaskConical, Settings, Sparkles, ChevronLeft } from 'lucide-react';

interface SidebarProps {
  collapsed: boolean;
  onToggleCollapse: () => void;
  activeView: string;
  onNavigate: (view: 'inicio' | 'preprocesamiento' | 'estadistica' | 'asistente') => void;
}

const navigationItems = [
  { name: 'Inicio', icon: LayoutDashboard, view: 'inicio' as const },
  { name: 'Preprocesamiento', icon: Database, view: 'preprocesamiento' as const },
  { name: 'Estadística Descriptiva', icon: BarChart3, view: 'estadistica' as const },
  { name: 'Asistente IA', icon: Sparkles, view: 'asistente' as const },
  { name: 'Pruebas de Hipótesis', icon: FlaskConical, view: null },
  { name: 'Configuración', icon: Settings, view: null },
];

export function Sidebar({ collapsed, onToggleCollapse, activeView, onNavigate }: SidebarProps) {
  return (
    <aside
      className={`
        relative bg-white border-r border-gray-200 shadow-xl z-20
        transition-all duration-300 ease-in-out flex flex-col
        ${collapsed ? 'w-0 overflow-hidden' : 'w-64'}
      `}
    >
      {/* Toggle Button — positioned at the edge, always visible */}
      <button
        onClick={onToggleCollapse}
        className={`
          absolute top-5 -right-3 z-30
          w-6 h-6 bg-white border border-gray-200 rounded-full
          shadow-md flex items-center justify-center
          hover:bg-teal-50 hover:border-teal-300
          transition-all duration-200
          ${collapsed ? 'opacity-0 pointer-events-none' : 'opacity-100'}
        `}
        aria-label={collapsed ? 'Abrir menú lateral' : 'Cerrar menú lateral'}
      >
        <ChevronLeft className={`w-3.5 h-3.5 text-gray-500 transition-transform duration-300`} />
      </button>

      {/* Inner content wrapper — prevents text wrapping during transition */}
      <div className={`
        flex flex-col h-full w-64 min-w-[16rem]
        transition-opacity duration-200 ease-in-out
        ${collapsed ? 'opacity-0' : 'opacity-100'}
      `}>
        {/* Logo Section */}
        <div className="h-16 flex items-center px-6 border-b border-gray-200">
          <div className="flex items-center gap-2.5">
            <div className="w-7 h-7 bg-gradient-to-br from-teal-600 to-teal-700 rounded flex items-center justify-center flex-shrink-0">
              <span className="text-white text-sm">B</span>
            </div>
            <span className="text-gray-900 tracking-tight whitespace-nowrap">Biometric</span>
          </div>
        </div>

        {/* Navigation */}
        <nav className="flex-1 py-6 px-3">
          <ul className="space-y-1">
            {navigationItems.map((item) => {
              const isActive = item.view === activeView;
              return (
                <li key={item.name}>
                  <button
                    onClick={() => item.view && onNavigate(item.view)}
                    disabled={!item.view}
                    className={`w-full flex items-center gap-3 px-3 py-2.5 rounded-md transition-all whitespace-nowrap ${isActive
                        ? 'bg-teal-50 text-teal-700'
                        : 'text-gray-700 hover:bg-gray-50 hover:text-gray-900'
                      } ${!item.view ? 'opacity-50 cursor-not-allowed' : ''}`}
                  >
                    <item.icon className="w-5 h-5 flex-shrink-0" />
                    <span className="text-sm">{item.name}</span>
                    {isActive && (
                      <div className="ml-auto w-1.5 h-1.5 bg-teal-600 rounded-full flex-shrink-0"></div>
                    )}
                  </button>
                </li>
              );
            })}
          </ul>
        </nav>

        {/* User Profile Section */}
        <div className="p-4 border-t border-gray-200">
          <div className="flex items-center gap-3">
            <div className="w-9 h-9 bg-gradient-to-br from-teal-100 to-teal-200 rounded-full flex items-center justify-center flex-shrink-0">
              <span className="text-teal-700 text-sm">SC</span>
            </div>
            <div className="flex-1 min-w-0">
              <div className="text-sm text-gray-900 truncate">Dr. Sara Castro</div>
              <div className="text-xs text-gray-500">Investigadora</div>
            </div>
          </div>
        </div>
      </div>
    </aside>
  );
}