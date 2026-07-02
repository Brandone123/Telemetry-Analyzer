import { Switch, Route, Router as WouterRouter } from "wouter";
import { QueryClient, QueryClientProvider } from "@tanstack/react-query";
import { Toaster } from "@/components/ui/toaster";
import { TooltipProvider } from "@/components/ui/tooltip";
import NotFound from "@/pages/not-found";
import Home from "@/pages/Home";
import NoComAnalysis from "@/pages/NoComAnalysis";
import RunningHoursAnalysis from "@/pages/RunningHoursAnalysis";
import FuelSensorAnalysis from "@/pages/FuelSensorAnalysis";
import ACAnalysis from "@/pages/ACAnalysis";

const queryClient = new QueryClient();

function Router() {
  return (
    <Switch>
      <Route path="/" component={Home} />
      <Route path="/analysis/no-com" component={NoComAnalysis} />
      <Route path="/analysis/running-hours" component={RunningHoursAnalysis} />
      <Route path="/analysis/fuel-sensor" component={FuelSensorAnalysis} />
      <Route path="/analysis/ac" component={ACAnalysis} />
      <Route component={NotFound} />
    </Switch>
  );
}

function App() {
  return (
    <QueryClientProvider client={queryClient}>
      <TooltipProvider>
        <WouterRouter base={import.meta.env.BASE_URL.replace(/\/$/, "")}>
          <Router />
        </WouterRouter>
        <Toaster />
      </TooltipProvider>
    </QueryClientProvider>
  );
}

export default App;
