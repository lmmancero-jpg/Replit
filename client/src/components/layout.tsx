import { SidebarProvider, SidebarTrigger } from "@/components/ui/sidebar";
import { AppSidebar } from "./app-sidebar";
import { ReactNode } from "react";

interface LayoutProps {
  children: ReactNode;
}

export function Layout({ children }: LayoutProps) {
  return (
    <SidebarProvider>
      <div className="flex h-screen w-full bg-background/50 overflow-hidden relative">
        <AppSidebar />
        
        <div className="flex flex-col flex-1 h-full max-h-screen relative z-10">
          <header className="h-16 flex items-center justify-between px-4 lg:px-8 border-b border-border/60 bg-background/80 backdrop-blur-md sticky top-0 z-20">
            <div className="flex items-center gap-4">
              <SidebarTrigger className="text-foreground/70 hover:text-foreground" />
              <div className="h-6 w-px bg-border mx-2 hidden md:block"></div>
              <h2 className="font-display font-semibold text-foreground hidden md:block">
                Centro de Operaciones
              </h2>
            </div>
            
            <div className="flex items-center gap-4">
              <div className="flex items-center gap-2 text-sm font-medium px-3 py-1.5 rounded-full bg-green-500/10 text-green-600 border border-green-500/20">
                <span className="relative flex h-2 w-2">
                  <span className="animate-ping absolute inline-flex h-full w-full rounded-full bg-green-400 opacity-75"></span>
                  <span className="relative inline-flex rounded-full h-2 w-2 bg-green-500"></span>
                </span>
                Sistema En Línea
              </div>
            </div>
          </header>
          
          <main className="flex-1 overflow-y-auto p-4 md:p-6 lg:p-8 scroll-smooth">
            <div className="max-w-[1600px] mx-auto">
              {children}
            </div>
          </main>
        </div>
      </div>
    </SidebarProvider>
  );
}
