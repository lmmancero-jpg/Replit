import { BarChart3, History, Settings, Zap } from "lucide-react";
import { Link, useLocation } from "wouter";
import {
  Sidebar,
  SidebarContent,
  SidebarGroup,
  SidebarGroupContent,
  SidebarGroupLabel,
  SidebarHeader,
  SidebarMenu,
  SidebarMenuButton,
  SidebarMenuItem,
  SidebarFooter,
} from "@/components/ui/sidebar";

export function AppSidebar() {
  const [location] = useLocation();

  const navItems = [
    { title: "Generador", url: "/", icon: Zap },
    { title: "Historial", url: "/history", icon: History },
    { title: "Métricas (Pronto)", url: "#", icon: BarChart3, disabled: true },
    { title: "Ajustes", url: "#", icon: Settings, disabled: true },
  ];

  return (
    <Sidebar className="border-r border-sidebar-border shadow-xl">
      <SidebarHeader className="p-6">
        <div className="flex items-center gap-3 font-display">
          <div className="w-8 h-8 bg-primary rounded-lg flex items-center justify-center shadow-lg shadow-primary/20">
            <Zap className="w-5 h-5 text-primary-foreground" />
          </div>
          <div>
            <h1 className="text-xl font-bold text-sidebar-foreground tracking-wide leading-none">
              NEXUS
            </h1>
            <p className="text-xs text-sidebar-foreground/60 uppercase tracking-widest font-medium">
              Power Plant Ops
            </p>
          </div>
        </div>
      </SidebarHeader>
      
      <SidebarContent>
        <SidebarGroup>
          <SidebarGroupLabel className="text-sidebar-foreground/50 text-xs font-semibold tracking-wider font-display px-6 py-2">
            INFORMES
          </SidebarGroupLabel>
          <SidebarGroupContent>
            <SidebarMenu className="px-3">
              {navItems.map((item) => {
                const isActive = location === item.url;
                
                return (
                  <SidebarMenuItem key={item.title}>
                    <SidebarMenuButton 
                      asChild 
                      isActive={isActive}
                      className={`
                        mb-1 h-10 px-3 rounded-md transition-all duration-200
                        ${isActive 
                          ? "bg-primary text-primary-foreground shadow-md font-medium" 
                          : "text-sidebar-foreground/80 hover:bg-sidebar-accent hover:text-sidebar-accent-foreground"}
                        ${item.disabled ? "opacity-50 pointer-events-none" : ""}
                      `}
                    >
                      <Link href={item.disabled ? "#" : item.url} className="flex items-center gap-3 w-full">
                        <item.icon className={`w-4 h-4 ${isActive ? "text-primary-foreground" : ""}`} />
                        <span>{item.title}</span>
                      </Link>
                    </SidebarMenuButton>
                  </SidebarMenuItem>
                );
              })}
            </SidebarMenu>
          </SidebarGroupContent>
        </SidebarGroup>
      </SidebarContent>

      <SidebarFooter className="p-6 text-xs text-sidebar-foreground/40 font-display text-center">
        IDOM System v6.0 <br/> © 2025
      </SidebarFooter>
    </Sidebar>
  );
}
