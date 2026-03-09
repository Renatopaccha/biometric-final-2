import { Bell, Menu, PanelLeftClose } from 'lucide-react';

interface HeaderProps {
  currentView: string;
  onToggleSidebar: () => void;
  isSidebarOpen: boolean;
}

export function Header({ currentView, onToggleSidebar, isSidebarOpen }: HeaderProps) {
  const titles = {
    inicio: {
      title: 'Inicio',
      subtitle: 'Panel de control y análisis'
    },
    preprocesamiento: {
      title: 'Preprocesamiento de Datos',
      subtitle: 'Importar, limpiar y preparar datos para análisis'
    },
    estadistica: {
      title: 'Estadística Descriptiva',
      subtitle: 'Análisis estadístico y visualización de datos'
    },
    asistente: {
      title: 'Asistente Virtual',
      subtitle: 'Chat inteligente para análisis de datos con IA'
    }
  };

  const current = titles[currentView as keyof typeof titles] || titles.inicio;

  return (
    <header className="h-16 bg-white border-b border-gray-200 flex items-center justify-between px-6">
      <div className="flex items-center gap-4">
        <button
          onClick={onToggleSidebar}
          className="p-2 hover:bg-gray-100 rounded-md transition-colors"
          aria-label={isSidebarOpen ? 'Cerrar menú lateral' : 'Abrir menú lateral'}
        >
          {isSidebarOpen ? (
            <PanelLeftClose className="w-5 h-5 text-gray-600" />
          ) : (
            <Menu className="w-5 h-5 text-gray-600" />
          )}
        </button>
        <div>
          <h1 className="text-gray-900">{current.title}</h1>
          <p className="text-xs text-gray-500">{current.subtitle}</p>
        </div>
      </div>

      <div className="flex items-center gap-3">
        <button className="relative p-2 text-gray-600 hover:bg-gray-50 rounded-md transition-colors">
          <Bell className="w-5 h-5" />
          <span className="absolute top-1.5 right-1.5 w-2 h-2 bg-teal-600 rounded-full"></span>
        </button>
      </div>
    </header>
  );
}