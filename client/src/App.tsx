import { Switch, Route } from "wouter";
import { queryClient } from "./lib/queryClient";
import { QueryClientProvider } from "@tanstack/react-query";
import { Toaster } from "@/components/ui/toaster";
import { TooltipProvider } from "@/components/ui/tooltip";
import { FileStoreProvider } from "@/lib/fileStore";
import NotFound from "@/pages/not-found";

import Generator from "./pages/generator";
import Metrics   from "./pages/metrics";

function Router() {
  return (
    <Switch>
      <Route path="/"        component={Generator} />
      <Route path="/metrics" component={Metrics} />
      <Route component={NotFound} />
    </Switch>
  );
}

function App() {
  return (
    <QueryClientProvider client={queryClient}>
      <TooltipProvider>
        <FileStoreProvider>
          <Toaster />
          <Router />
        </FileStoreProvider>
      </TooltipProvider>
    </QueryClientProvider>
  );
}

export default App;
