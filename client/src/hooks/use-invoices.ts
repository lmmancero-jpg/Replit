import { useQuery, useMutation } from "@tanstack/react-query";
import { queryClient } from "@/lib/queryClient";
import { apiRequest } from "@/lib/queryClient";
import type { Invoice, InsertInvoice } from "@shared/schema";

const INVOICES_KEY = "/api/invoices";

function invoicesKey(period?: string) {
  return period ? [INVOICES_KEY, period] : [INVOICES_KEY];
}

export function useInvoices(period?: string) {
  return useQuery<Invoice[]>({
    queryKey: invoicesKey(period),
    queryFn: async () => {
      const url = period ? `${INVOICES_KEY}?period=${encodeURIComponent(period)}` : INVOICES_KEY;
      const res = await fetch(url);
      if (!res.ok) throw new Error(`Error ${res.status}`);
      return res.json();
    },
  });
}

export function useInvoiceSummary(period: string) {
  return useQuery<Record<string, number>>({
    queryKey: ["/api/invoices/summary", period],
    queryFn: async () => {
      const res = await fetch(`/api/invoices/summary?period=${encodeURIComponent(period)}`);
      if (!res.ok) throw new Error(`Error ${res.status}`);
      return res.json();
    },
    enabled: !!period,
  });
}

export function useCreateInvoice(period?: string) {
  return useMutation({
    mutationFn: (data: Omit<InsertInvoice, "id">) =>
      apiRequest("POST", INVOICES_KEY, data),
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: invoicesKey(period) });
      queryClient.invalidateQueries({ queryKey: ["/api/invoices/summary", period] });
    },
  });
}

export function useUpdateInvoice(period?: string) {
  return useMutation({
    mutationFn: ({ id, data }: { id: number; data: Partial<InsertInvoice> }) =>
      apiRequest("PUT", `${INVOICES_KEY}/${id}`, data),
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: invoicesKey(period) });
      queryClient.invalidateQueries({ queryKey: ["/api/invoices/summary", period] });
    },
  });
}

export function useDeleteInvoice(period?: string) {
  return useMutation({
    mutationFn: (id: number) => apiRequest("DELETE", `${INVOICES_KEY}/${id}`),
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: invoicesKey(period) });
      queryClient.invalidateQueries({ queryKey: ["/api/invoices/summary", period] });
    },
  });
}
