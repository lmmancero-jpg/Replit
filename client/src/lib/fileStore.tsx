import { createContext, useContext, useState, useCallback, useEffect, ReactNode } from "react";
import * as XLSX from "xlsx";

interface FileStoreState {
  prodFile: File | null;
  aforoFile: File | null;
  wbProd: XLSX.WorkBook | null;
  wbAforo: XLSX.WorkBook | null;
  fileNameProd: string;
  fileNameAforo: string;
  prodLoading: boolean;
  aforoLoading: boolean;
  setProdEntry: (file: File) => void;
  setAforoEntry: (file: File) => void;
}

const FileStore = createContext<FileStoreState | null>(null);

export function FileStoreProvider({ children }: { children: ReactNode }) {
  const [prodFile, setProdFile] = useState<File | null>(null);
  const [aforoFile, setAforoFile] = useState<File | null>(null);
  const [wbProd, setWbProd] = useState<XLSX.WorkBook | null>(null);
  const [wbAforo, setWbAforo] = useState<XLSX.WorkBook | null>(null);
  const [fileNameProd, setFileNameProd] = useState("");
  const [fileNameAforo, setFileNameAforo] = useState("");
  const [prodLoading, setProdLoading] = useState(false);
  const [aforoLoading, setAforoLoading] = useState(false);

  const setProdEntry = useCallback((file: File) => {
    setProdFile(file);
    setFileNameProd(file.name);
    setProdLoading(true);
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const wb = XLSX.read(e.target!.result as ArrayBuffer, { type: "array", cellDates: true });
        setWbProd(wb);
      } catch {
        setWbProd(null);
      } finally {
        setProdLoading(false);
      }
    };
    reader.onerror = () => { setWbProd(null); setProdLoading(false); };
    reader.readAsArrayBuffer(file);
  }, []);

  const setAforoEntry = useCallback((file: File) => {
    setAforoFile(file);
    setFileNameAforo(file.name);
    setAforoLoading(true);
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const wb = XLSX.read(e.target!.result as ArrayBuffer, { type: "array", cellDates: true });
        setWbAforo(wb);
      } catch {
        setWbAforo(null);
      } finally {
        setAforoLoading(false);
      }
    };
    reader.onerror = () => { setWbAforo(null); setAforoLoading(false); };
    reader.readAsArrayBuffer(file);
  }, []);

  return (
    <FileStore.Provider value={{
      prodFile, aforoFile, wbProd, wbAforo,
      fileNameProd, fileNameAforo,
      prodLoading, aforoLoading,
      setProdEntry, setAforoEntry,
    }}>
      {children}
    </FileStore.Provider>
  );
}

export function useFileStore() {
  const ctx = useContext(FileStore);
  if (!ctx) throw new Error("useFileStore must be used within FileStoreProvider");
  return ctx;
}

export { useEffect };
