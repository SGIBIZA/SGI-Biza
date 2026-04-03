/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useEffect, useMemo } from 'react';
import { 
  auth, 
  db, 
  googleProvider, 
  signInWithPopup, 
  signOut, 
  onAuthStateChanged, 
  collection, 
  doc, 
  getDoc, 
  getDocs, 
  setDoc, 
  updateDoc, 
  deleteDoc, 
  query, 
  where, 
  orderBy, 
  onSnapshot, 
  Timestamp, 
  serverTimestamp, 
  runTransaction,
  FirebaseUser,
  OperationType,
  handleFirestoreError
} from './firebase';
import { 
  format, 
  addMinutes, 
  isBefore, 
  isAfter, 
  parse, 
  startOfDay, 
  differenceInDays, 
  addDays, 
  isSameDay, 
  subHours,
  parseISO
} from 'date-fns';
import * as XLSX from 'xlsx';
import { 
  Calendar, 
  Clock, 
  User, 
  CheckCircle, 
  XCircle, 
  AlertCircle, 
  LogOut, 
  Settings, 
  FileUp, 
  Users, 
  ListOrdered, 
  ChevronRight, 
  ChevronLeft,
  Trash2,
  Check,
  X,
  History
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';
import { clsx, type ClassValue } from 'clsx';
import { twMerge } from 'tailwind-merge';

function cn(...inputs: ClassValue[]) {
  return twMerge(clsx(inputs));
}

// --- Types ---

interface UserProfile {
  id: string;
  name: string;
  email: string;
  sector?: string;
  role: 'admin' | 'user';
  total_appointments: number;
  no_shows: number;
  normal_cancellations: number;
  late_cancellations: number;
  rejections: number;
  last_appointment_date?: string;
  priority_score: number;
  created_at: string;
}

interface Appointment {
  id: string;
  user_id: string;
  user_name: string;
  date: string; // YYYY-MM-DD
  time_slot: string; // HH:mm
  status: 'pending' | 'approved' | 'rejected' | 'cancelled' | 'no-show' | 'completed';
  created_at: string;
  approved_by?: string;
  cancelled_at?: string;
  cancellation_type?: 'normal' | 'late';
  imported?: boolean;
}

interface WaitlistEntry {
  id: string;
  user_id: string;
  user_name: string;
  date: string;
  priority_score: number;
  joined_at: string;
}

interface AppSettings {
  service_day: number; // 0-6
  release_time: string; // HH:mm
  open_day_before: boolean;
  max_appointments_per_period: number;
}

// --- Constants ---

const TIME_SLOTS = [
  '10:00', '10:30', '11:00', '11:30', '12:00', '12:30',
  '14:00', '14:30', '15:00', '15:30', '16:00', '16:30', '17:00', '17:30'
];

const STATUS_LABELS: Record<string, { label: string; color: string }> = {
  pending: { label: 'Pendente', color: 'bg-yellow-100 text-yellow-800 border-yellow-200' },
  approved: { label: 'Aprovado', color: 'bg-blue-100 text-blue-800 border-blue-200' },
  rejected: { label: 'Rejeitado', color: 'bg-red-100 text-red-800 border-red-200' },
  cancelled: { label: 'Cancelado', color: 'bg-gray-100 text-gray-800 border-gray-200' },
  'no-show': { label: 'Falta', color: 'bg-orange-100 text-orange-800 border-orange-200' },
  completed: { label: 'Realizado', color: 'bg-green-100 text-green-800 border-green-200' },
};

// --- Error Boundary ---

class ErrorBoundary extends React.Component<{ children: React.ReactNode }, { hasError: boolean; errorInfo: string | null }> {
  constructor(props: { children: React.ReactNode }) {
    super(props);
    this.state = { hasError: false, errorInfo: null };
  }

  static getDerivedStateFromError(error: any) {
    return { hasError: true, errorInfo: error.message };
  }

  componentDidCatch(error: any, errorInfo: any) {
    console.error("ErrorBoundary caught an error", error, errorInfo);
  }

  render() {
    if (this.state.hasError) {
      return (
        <div className="min-h-screen flex items-center justify-center bg-gray-50 p-4">
          <div className="max-w-md w-full bg-white rounded-2xl shadow-xl p-8 text-center">
            <div className="w-16 h-16 bg-red-100 rounded-full flex items-center justify-center mx-auto mb-6">
              <AlertCircle className="w-8 h-8 text-red-600" />
            </div>
            <h2 className="text-2xl font-bold text-gray-900 mb-2">Ops! Algo deu errado.</h2>
            <p className="text-gray-600 mb-6">Ocorreu um erro inesperado. Por favor, tente recarregar a página.</p>
            <div className="bg-gray-50 p-4 rounded-lg text-left mb-6 overflow-auto max-h-40">
              <code className="text-xs text-red-500">{this.state.errorInfo}</code>
            </div>
            <button
              onClick={() => window.location.reload()}
              className="w-full bg-blue-600 text-white py-3 rounded-xl font-medium hover:bg-blue-700 transition-colors"
            >
              Recarregar Página
            </button>
          </div>
        </div>
      );
    }

    return this.props.children;
  }
}

// --- Components ---

export default function App() {
  return (
    <ErrorBoundary>
      <AppContent />
    </ErrorBoundary>
  );
}

function AppContent() {
  const [user, setUser] = useState<FirebaseUser | null>(null);
  const [profile, setProfile] = useState<UserProfile | null>(null);
  const [settings, setSettings] = useState<AppSettings | null>(null);
  const [loading, setLoading] = useState(true);
  const [view, setView] = useState<'agenda' | 'admin' | 'history'>('agenda');

  useEffect(() => {
    const unsubscribe = onAuthStateChanged(auth, async (firebaseUser) => {
      if (firebaseUser) {
        setUser(firebaseUser);
        try {
          const profileRef = doc(db, 'users', firebaseUser.uid);
          const profileSnap = await getDoc(profileRef);
          
          if (profileSnap.exists()) {
            setProfile(profileSnap.data() as UserProfile);
          } else {
            const newProfile: UserProfile = {
              id: firebaseUser.uid,
              name: firebaseUser.displayName || 'Colaborador',
              email: firebaseUser.email || '',
              role: firebaseUser.email === 'ghtbizasgi@gmail.com' ? 'admin' : 'user',
              total_appointments: 0,
              no_shows: 0,
              normal_cancellations: 0,
              late_cancellations: 0,
              rejections: 0,
              priority_score: 100,
              created_at: new Date().toISOString(),
            };
            await setDoc(profileRef, newProfile);
            setProfile(newProfile);
          }

          // Fetch settings
          const settingsRef = doc(db, 'settings', 'config');
          const settingsSnap = await getDoc(settingsRef);
          if (settingsSnap.exists()) {
            setSettings(settingsSnap.data() as AppSettings);
          } else {
            const defaultSettings: AppSettings = {
              service_day: 4, // Thursday
              release_time: '08:00',
              open_day_before: true,
              max_appointments_per_period: 14
            };
            await setDoc(settingsRef, defaultSettings);
            setSettings(defaultSettings);
          }
        } catch (error) {
          console.error('Error loading profile or settings:', error);
        }
      } else {
        setUser(null);
        setProfile(null);
        setSettings(null);
      }
      setLoading(false);
    });

    return () => unsubscribe();
  }, []);

  const handleLogin = async () => {
    try {
      await signInWithPopup(auth, googleProvider);
    } catch (error) {
      console.error('Login error:', error);
    }
  };

  const handleLogout = async () => {
    try {
      await signOut(auth);
    } catch (error) {
      console.error('Logout error:', error);
    }
  };

  if (loading || (user && !profile) || (user && !settings)) {
    return (
      <div className="min-h-screen flex items-center justify-center bg-gray-50">
        <div className="animate-spin rounded-full h-12 w-12 border-t-2 border-b-2 border-blue-600"></div>
      </div>
    );
  }

  if (!user) {
    return (
      <div className="min-h-screen flex flex-col items-center justify-center bg-gray-50 p-4">
        <motion.div 
          initial={{ opacity: 0, y: 20 }}
          animate={{ opacity: 1, y: 0 }}
          className="max-w-md w-full bg-white rounded-2xl shadow-xl p-8 text-center"
        >
          <div className="w-20 h-20 bg-blue-100 rounded-full flex items-center justify-center mx-auto mb-6">
            <Calendar className="w-10 h-10 text-blue-600" />
          </div>
          <h1 className="text-3xl font-bold text-gray-900 mb-2">Massoterapia</h1>
          <p className="text-gray-600 mb-8">Gestão de agendamentos para colaboradores da metalúrgica.</p>
          <button
            onClick={handleLogin}
            className="w-full flex items-center justify-center gap-3 bg-white border border-gray-300 text-gray-700 px-6 py-3 rounded-xl font-medium hover:bg-gray-50 transition-colors shadow-sm"
          >
            <img src="https://www.gstatic.com/firebasejs/ui/2.0.0/images/auth/google.svg" alt="Google" className="w-5 h-5" />
            Entrar com Google
          </button>
        </motion.div>
      </div>
    );
  }

  return (
    <div className="min-h-screen bg-gray-50 flex flex-col">
      {/* Header */}
      <header className="bg-white border-b border-gray-200 sticky top-0 z-10">
        <div className="max-w-7xl mx-auto px-4 h-16 flex items-center justify-between">
          <div className="flex items-center gap-2">
            <div className="w-8 h-8 bg-blue-600 rounded-lg flex items-center justify-center">
              <Calendar className="w-5 h-5 text-white" />
            </div>
            <span className="font-bold text-xl text-gray-900 hidden sm:block">MassoManager</span>
          </div>

          <nav className="flex items-center gap-1 sm:gap-4">
            <button 
              onClick={() => setView('agenda')}
              className={cn(
                "px-3 py-2 rounded-lg text-sm font-medium transition-colors",
                view === 'agenda' ? "bg-blue-50 text-blue-700" : "text-gray-600 hover:bg-gray-100"
              )}
            >
              Agenda
            </button>
            <button 
              onClick={() => setView('history')}
              className={cn(
                "px-3 py-2 rounded-lg text-sm font-medium transition-colors",
                view === 'history' ? "bg-blue-50 text-blue-700" : "text-gray-600 hover:bg-gray-100"
              )}
            >
              Meu Histórico
            </button>
            {profile?.role === 'admin' && (
              <button 
                onClick={() => setView('admin')}
                className={cn(
                  "px-3 py-2 rounded-lg text-sm font-medium transition-colors",
                  view === 'admin' ? "bg-blue-50 text-blue-700" : "text-gray-600 hover:bg-gray-100"
                )}
              >
                Painel Admin
              </button>
            )}
          </nav>

          <div className="flex items-center gap-3">
            <div className="hidden md:flex flex-col items-end">
              <span className="text-sm font-medium text-gray-900">{profile?.name}</span>
              <span className="text-xs text-gray-500">Score: {profile?.priority_score.toFixed(0)}</span>
            </div>
            <button 
              onClick={handleLogout}
              className="p-2 text-gray-400 hover:text-red-600 hover:bg-red-50 rounded-lg transition-colors"
              title="Sair"
            >
              <LogOut className="w-5 h-5" />
            </button>
          </div>
        </div>
      </header>

      {/* Main Content */}
      <main className="flex-1 max-w-7xl w-full mx-auto p-4 sm:p-6">
        <AnimatePresence mode="wait">
          {view === 'agenda' && <AgendaView profile={profile!} settings={settings!} />}
          {view === 'admin' && <AdminDashboard profile={profile!} settings={settings!} setSettings={setSettings} />}
          {view === 'history' && <UserHistory profile={profile!} />}
        </AnimatePresence>
      </main>
    </div>
  );
}

// --- View Components ---

function AgendaView({ profile, settings }: { profile: UserProfile; settings: AppSettings }) {
  const [selectedDate, setSelectedDate] = useState(format(new Date(), 'yyyy-MM-dd'));
  const [appointments, setAppointments] = useState<Appointment[]>([]);
  const [waitlist, setWaitlist] = useState<WaitlistEntry[]>([]);
  const [loading, setLoading] = useState(true);

  useEffect(() => {
    const q = query(collection(db, 'appointments'), where('date', '==', selectedDate));
    const unsubscribe = onSnapshot(q, (snapshot) => {
      const apps = snapshot.docs.map(doc => ({ id: doc.id, ...doc.data() } as Appointment));
      setAppointments(apps);
      setLoading(false);
    }, (error) => {
      handleFirestoreError(error, OperationType.LIST, 'appointments');
    });

    return () => unsubscribe();
  }, [selectedDate]);

  useEffect(() => {
    const q = query(collection(db, 'waitlist'), where('date', '==', selectedDate), orderBy('priority_score', 'desc'));
    const unsubscribe = onSnapshot(q, (snapshot) => {
      const wait = snapshot.docs.map(doc => ({ id: doc.id, ...doc.data() } as WaitlistEntry));
      setWaitlist(wait);
    }, (error) => {
      handleFirestoreError(error, OperationType.LIST, 'waitlist');
    });

    return () => unsubscribe();
  }, [selectedDate]);

  const isAgendaOpen = useMemo(() => {
    if (profile.role === 'admin') return true;
    
    const date = parseISO(selectedDate);
    if (date.getDay() !== settings.service_day) return false;

    const now = new Date();
    const releaseDay = settings.open_day_before ? addDays(date, -1) : date;
    const releaseDateTime = parse(`${format(releaseDay, 'yyyy-MM-dd')} ${settings.release_time}`, 'yyyy-MM-dd HH:mm', new Date());

    return isAfter(now, releaseDateTime);
  }, [selectedDate, settings, profile.role]);

  const handleBook = async (slot: string) => {
    if (!isAgendaOpen) {
      alert('A agenda para este dia ainda não está aberta.');
      return;
    }
    if (appointments.some(a => a.time_slot === slot && ['approved', 'pending'].includes(a.status))) {
      return;
    }

    const userHasAppOnDay = appointments.some(a => a.user_id === profile.id && ['approved', 'pending'].includes(a.status));
    if (userHasAppOnDay) {
      alert('Você já possui um agendamento para este dia.');
      return;
    }

    try {
      const newApp: Omit<Appointment, 'id'> = {
        user_id: profile.id,
        user_name: profile.name,
        date: selectedDate,
        time_slot: slot,
        status: 'pending',
        created_at: new Date().toISOString(),
      };
      await setDoc(doc(collection(db, 'appointments')), newApp);
    } catch (error) {
      handleFirestoreError(error, OperationType.CREATE, 'appointments');
    }
  };

  const handleCancel = async (appId: string, slotTime: string) => {
    if (!window.confirm('Tem certeza que deseja cancelar seu agendamento?')) return;

    try {
      const appRef = doc(db, 'appointments', appId);
      const now = new Date();
      const appDateTime = parse(`${selectedDate} ${slotTime}`, 'yyyy-MM-dd HH:mm', new Date());
      
      const isLate = differenceInDays(appDateTime, now) === 0 && isBefore(subHours(appDateTime, 1), now);
      
      await updateDoc(appRef, {
        status: 'cancelled',
        cancelled_at: now.toISOString(),
        cancellation_type: isLate ? 'late' : 'normal'
      });

      // Update user stats
      const userRef = doc(db, 'users', profile.id);
      await updateDoc(userRef, {
        [isLate ? 'late_cancellations' : 'normal_cancellations']: (profile[isLate ? 'late_cancellations' : 'normal_cancellations'] || 0) + 1,
        priority_score: calculateScore({
          ...profile,
          [isLate ? 'late_cancellations' : 'normal_cancellations']: (profile[isLate ? 'late_cancellations' : 'normal_cancellations'] || 0) + 1
        })
      });

    } catch (error) {
      handleFirestoreError(error, OperationType.UPDATE, 'appointments');
    }
  };

  const handleJoinWaitlist = async () => {
    if (waitlist.some(w => w.user_id === profile.id)) return;

    try {
      const newEntry: Omit<WaitlistEntry, 'id'> = {
        user_id: profile.id,
        user_name: profile.name,
        date: selectedDate,
        priority_score: profile.priority_score,
        joined_at: new Date().toISOString(),
      };
      await setDoc(doc(collection(db, 'waitlist')), newEntry);
    } catch (error) {
      handleFirestoreError(error, OperationType.CREATE, 'waitlist');
    }
  };

  const handleLeaveWaitlist = async () => {
    const entry = waitlist.find(w => w.user_id === profile.id);
    if (!entry) return;

    try {
      await deleteDoc(doc(db, 'waitlist', entry.id));
    } catch (error) {
      handleFirestoreError(error, OperationType.DELETE, 'waitlist');
    }
  };

  const isWaitlisted = waitlist.some(w => w.user_id === profile.id);
  const myAppointment = appointments.find(a => a.user_id === profile.id && ['approved', 'pending'].includes(a.status));

  return (
    <motion.div 
      initial={{ opacity: 0 }}
      animate={{ opacity: 1 }}
      exit={{ opacity: 0 }}
      className="space-y-6"
    >
      <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4">
        <div>
          <h2 className="text-2xl font-bold text-gray-900">Agenda de Massoterapia</h2>
          <p className="text-gray-500">Selecione um horário para agendar seu benefício.</p>
        </div>
        <div className="flex items-center gap-2 bg-white p-1 rounded-xl border border-gray-200 shadow-sm">
          <button 
            onClick={() => setSelectedDate(format(addDays(parseISO(selectedDate), -1), 'yyyy-MM-dd'))}
            className="p-2 hover:bg-gray-100 rounded-lg transition-colors"
          >
            <ChevronLeft className="w-5 h-5" />
          </button>
          <div className="px-4 font-medium text-gray-900">
            {format(parseISO(selectedDate), "dd 'de' MMMM", { locale: undefined })}
          </div>
          <button 
            onClick={() => setSelectedDate(format(addDays(parseISO(selectedDate), 1), 'yyyy-MM-dd'))}
            className="p-2 hover:bg-gray-100 rounded-lg transition-colors"
          >
            <ChevronRight className="w-5 h-5" />
          </button>
        </div>
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
        {/* Slots Grid */}
        <div className="lg:col-span-2 space-y-4">
          {!isAgendaOpen && parseISO(selectedDate).getDay() === settings.service_day && (
            <div className="bg-blue-50 border border-blue-100 p-4 rounded-xl flex items-center gap-3 text-blue-800">
              <Clock className="w-5 h-5" />
              <p className="text-sm font-medium">
                A agenda para este dia será aberta em {settings.open_day_before ? 'um dia antes' : 'no próprio dia'} às {settings.release_time}.
              </p>
            </div>
          )}
          
          {parseISO(selectedDate).getDay() !== settings.service_day && (
            <div className="bg-gray-50 border border-gray-100 p-4 rounded-xl flex items-center gap-3 text-gray-500">
              <AlertCircle className="w-5 h-5" />
              <p className="text-sm font-medium">Não há atendimentos previstos para este dia.</p>
            </div>
          )}

          <div className="grid grid-cols-2 sm:grid-cols-3 gap-3">
            {TIME_SLOTS.map((slot) => {
              const app = appointments.find(a => a.time_slot === slot && ['approved', 'pending'].includes(a.status));
              const isOccupied = !!app;
              const isMine = app?.user_id === profile.id;

              return (
                <button
                  key={slot}
                  disabled={(isOccupied && !isMine) || (!isAgendaOpen && !isMine)}
                  onClick={() => isMine ? handleCancel(app.id, slot) : handleBook(slot)}
                  className={cn(
                    "p-4 rounded-xl border-2 transition-all flex flex-col items-center gap-2",
                    isMine 
                      ? "bg-blue-50 border-blue-500 text-blue-700" 
                      : isOccupied 
                        ? "bg-gray-50 border-gray-100 text-gray-400 cursor-not-allowed" 
                        : !isAgendaOpen
                          ? "bg-gray-50 border-gray-100 text-gray-400 cursor-not-allowed opacity-60"
                          : "bg-white border-gray-100 hover:border-blue-200 hover:bg-blue-50/30 text-gray-700"
                  )}
                >
                  <Clock className={cn("w-5 h-5", isMine ? "text-blue-600" : "text-gray-400")} />
                  <span className="font-bold text-lg">{slot}</span>
                  <span className="text-xs uppercase tracking-wider font-semibold">
                    {isMine ? "Seu Horário" : isOccupied ? "Ocupado" : !isAgendaOpen ? "Fechado" : "Disponível"}
                  </span>
                </button>
              );
            })}
          </div>
        </div>

        {/* Sidebar: Waitlist & Status */}
        <div className="space-y-6">
          {/* My Status Card */}
          <div className="bg-white rounded-2xl border border-gray-200 p-6 shadow-sm">
            <h3 className="font-bold text-gray-900 mb-4 flex items-center gap-2">
              <User className="w-5 h-5 text-blue-600" />
              Seu Status
            </h3>
            {myAppointment ? (
              <div className="space-y-4">
                <div className={cn("p-4 rounded-xl border flex flex-col gap-1", STATUS_LABELS[myAppointment.status].color)}>
                  <span className="text-xs font-bold uppercase opacity-70">Agendamento</span>
                  <span className="text-lg font-bold">{myAppointment.time_slot}</span>
                  <span className="text-sm">{STATUS_LABELS[myAppointment.status].label}</span>
                </div>
                <button 
                  onClick={() => handleCancel(myAppointment.id, myAppointment.time_slot)}
                  className="w-full py-2 text-sm font-medium text-red-600 hover:bg-red-50 rounded-lg transition-colors border border-red-100"
                >
                  Cancelar Agendamento
                </button>
              </div>
            ) : (
              <div className="text-center py-4">
                <p className="text-gray-500 text-sm mb-4">Você não tem agendamentos para hoje.</p>
                {isWaitlisted ? (
                  <div className="space-y-3">
                    <div className="p-3 bg-yellow-50 border border-yellow-100 rounded-xl text-yellow-800 text-sm">
                      Você está na fila de espera.
                    </div>
                    <button 
                      onClick={handleLeaveWaitlist}
                      className="text-sm text-gray-500 hover:text-red-600 underline"
                    >
                      Sair da fila
                    </button>
                  </div>
                ) : (
                  <button 
                    onClick={handleJoinWaitlist}
                    className="w-full py-3 bg-gray-900 text-white rounded-xl font-medium hover:bg-gray-800 transition-colors"
                  >
                    Entrar na Fila de Espera
                  </button>
                )}
              </div>
            )}
          </div>

          {/* Waitlist Card */}
          <div className="bg-white rounded-2xl border border-gray-200 p-6 shadow-sm">
            <h3 className="font-bold text-gray-900 mb-4 flex items-center gap-2">
              <ListOrdered className="w-5 h-5 text-blue-600" />
              Fila de Espera ({waitlist.length})
            </h3>
            <div className="space-y-3">
              {waitlist.length > 0 ? (
                waitlist.map((entry, idx) => (
                  <div key={entry.id} className="flex items-center justify-between p-3 bg-gray-50 rounded-xl border border-gray-100">
                    <div className="flex items-center gap-3">
                      <span className="text-xs font-bold text-gray-400 w-4">{idx + 1}</span>
                      <span className="text-sm font-medium text-gray-700">{entry.user_name}</span>
                    </div>
                    <span className="text-xs font-bold text-blue-600 bg-blue-50 px-2 py-1 rounded-md">
                      {entry.priority_score.toFixed(0)}
                    </span>
                  </div>
                ))
              ) : (
                <p className="text-center text-gray-400 text-sm py-4">Fila vazia</p>
              )}
            </div>
          </div>
        </div>
      </div>
    </motion.div>
  );
}

function AdminDashboard({ profile, settings, setSettings }: { profile: UserProfile; settings: AppSettings; setSettings: (s: AppSettings) => void }) {
  const [appointments, setAppointments] = useState<Appointment[]>([]);
  const [waitlist, setWaitlist] = useState<WaitlistEntry[]>([]);
  const [users, setUsers] = useState<UserProfile[]>([]);
  const [selectedDate, setSelectedDate] = useState(format(new Date(), 'yyyy-MM-dd'));
  const [activeTab, setActiveTab] = useState<'approvals' | 'users' | 'import' | 'settings'>('approvals');
  const [localSettings, setLocalSettings] = useState<AppSettings>(settings);

  useEffect(() => {
    setLocalSettings(settings);
  }, [settings]);

  const handleSaveSettings = async () => {
    try {
      await setDoc(doc(db, 'settings', 'config'), localSettings);
      setSettings(localSettings);
      alert('Configurações salvas com sucesso!');
    } catch (error) {
      handleFirestoreError(error, OperationType.UPDATE, 'settings');
    }
  };

  useEffect(() => {
    const q = query(collection(db, 'appointments'), where('date', '==', selectedDate), orderBy('created_at', 'desc'));
    const unsubscribe = onSnapshot(q, (snapshot) => {
      setAppointments(snapshot.docs.map(doc => ({ id: doc.id, ...doc.data() } as Appointment)));
    }, (error) => {
      handleFirestoreError(error, OperationType.LIST, 'appointments');
    });
    return () => unsubscribe();
  }, [selectedDate]);

  useEffect(() => {
    const q = query(collection(db, 'waitlist'), where('date', '==', selectedDate), orderBy('priority_score', 'desc'));
    const unsubscribe = onSnapshot(q, (snapshot) => {
      setWaitlist(snapshot.docs.map(doc => ({ id: doc.id, ...doc.data() } as WaitlistEntry)));
    }, (error) => {
      handleFirestoreError(error, OperationType.LIST, 'waitlist');
    });
    return () => unsubscribe();
  }, [selectedDate]);

  useEffect(() => {
    const q = query(collection(db, 'users'), orderBy('priority_score', 'desc'));
    const unsubscribe = onSnapshot(q, (snapshot) => {
      setUsers(snapshot.docs.map(doc => ({ id: doc.id, ...doc.data() } as UserProfile)));
    }, (error) => {
      handleFirestoreError(error, OperationType.LIST, 'users');
    });
    return () => unsubscribe();
  }, []);

  const handleStatusChange = async (appId: string, newStatus: Appointment['status']) => {
    try {
      const appRef = doc(db, 'appointments', appId);
      const appSnap = await getDoc(appRef);
      if (!appSnap.exists()) return;
      const app = appSnap.data() as Appointment;

      await updateDoc(appRef, { 
        status: newStatus,
        approved_by: profile.id
      });

      // Update user stats if completed or no-show
      if (newStatus === 'completed' || newStatus === 'no-show' || newStatus === 'rejected') {
        const userRef = doc(db, 'users', app.user_id);
        const userSnap = await getDoc(userRef);
        if (userSnap.exists()) {
          const userData = userSnap.data() as UserProfile;
          const updates: Partial<UserProfile> = {};
          if (newStatus === 'completed') {
            updates.total_appointments = (userData.total_appointments || 0) + 1;
            updates.last_appointment_date = new Date().toISOString();
          } else if (newStatus === 'no-show') {
            updates.no_shows = (userData.no_shows || 0) + 1;
          } else if (newStatus === 'rejected') {
            updates.rejections = (userData.rejections || 0) + 1;
          }
          updates.priority_score = calculateScore({ ...userData, ...updates });
          await updateDoc(userRef, updates);
        }
      }
    } catch (error) {
      handleFirestoreError(error, OperationType.UPDATE, 'appointments');
    }
  };

  const handleImport = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = async (evt) => {
      const bstr = evt.target?.result;
      const wb = XLSX.read(bstr, { type: 'binary' });
      const wsname = wb.SheetNames[0];
      const ws = wb.Sheets[wsname];
      const data = XLSX.utils.sheet_to_json(ws) as any[];

      for (const row of data) {
        // Simple logic: find user by name or create
        const userName = row['Colaborador'];
        if (!userName) continue;

        const userQ = query(collection(db, 'users'), where('name', '==', userName));
        const userSnap = await getDocs(userQ);
        let userId = '';

        if (userSnap.empty) {
          const newUserRef = doc(collection(db, 'users'));
          userId = newUserRef.id;
          const newUser: UserProfile = {
            id: userId,
            name: userName,
            email: `${userName.toLowerCase().replace(/\s/g, '.')}@empresa.com`,
            role: 'user',
            total_appointments: 1,
            no_shows: 0,
            normal_cancellations: 0,
            late_cancellations: 0,
            rejections: 0,
            last_appointment_date: new Date().toISOString(),
            priority_score: 100,
            created_at: new Date().toISOString(),
          };
          newUser.priority_score = calculateScore(newUser);
          await setDoc(newUserRef, newUser);
        } else {
          userId = userSnap.docs[0].id;
          const userData = userSnap.docs[0].data() as UserProfile;
          const updates: Partial<UserProfile> = {
            total_appointments: (userData.total_appointments || 0) + 1,
            last_appointment_date: new Date().toISOString(),
          };
          updates.priority_score = calculateScore({ ...userData, ...updates });
          await updateDoc(doc(db, 'users', userId), updates);
        }

        // Create completed appointment
        await setDoc(doc(collection(db, 'appointments')), {
          user_id: userId,
          user_name: userName,
          date: format(new Date(), 'yyyy-MM-dd'), // Placeholder date
          time_slot: '00:00',
          status: 'completed',
          created_at: new Date().toISOString(),
          imported: true
        });
      }
      alert('Importação concluída!');
    };
    reader.readAsBinaryString(file);
  };

  return (
    <motion.div 
      initial={{ opacity: 0 }}
      animate={{ opacity: 1 }}
      exit={{ opacity: 0 }}
      className="space-y-6"
    >
      <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4">
        <h2 className="text-2xl font-bold text-gray-900">Painel Administrativo</h2>
        <div className="flex bg-white p-1 rounded-xl border border-gray-200 shadow-sm">
          <button 
            onClick={() => setActiveTab('approvals')}
            className={cn("px-4 py-2 rounded-lg text-sm font-medium transition-colors", activeTab === 'approvals' ? "bg-blue-600 text-white" : "text-gray-600 hover:bg-gray-50")}
          >
            Aprovações
          </button>
          <button 
            onClick={() => setActiveTab('users')}
            className={cn("px-4 py-2 rounded-lg text-sm font-medium transition-colors", activeTab === 'users' ? "bg-blue-600 text-white" : "text-gray-600 hover:bg-gray-50")}
          >
            Colaboradores
          </button>
          <button 
            onClick={() => setActiveTab('import')}
            className={cn("px-4 py-2 rounded-lg text-sm font-medium transition-colors", activeTab === 'import' ? "bg-blue-600 text-white" : "text-gray-600 hover:bg-gray-50")}
          >
            Importar
          </button>
          <button 
            onClick={() => setActiveTab('settings')}
            className={cn("px-4 py-2 rounded-lg text-sm font-medium transition-colors", activeTab === 'settings' ? "bg-blue-600 text-white" : "text-gray-600 hover:bg-gray-50")}
          >
            Configurações
          </button>
        </div>
      </div>

      {activeTab === 'approvals' && (
        <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
          <div className="lg:col-span-2 space-y-4">
            <div className="flex items-center justify-between mb-2">
              <h3 className="font-bold text-gray-900">Solicitações do Dia</h3>
              <input 
                type="date" 
                value={selectedDate} 
                onChange={(e) => setSelectedDate(e.target.value)}
                className="text-sm border border-gray-200 rounded-lg p-2"
              />
            </div>
            <div className="bg-white rounded-2xl border border-gray-200 overflow-hidden shadow-sm">
              <table className="w-full text-left border-collapse">
                <thead className="bg-gray-50 border-b border-gray-200">
                  <tr>
                    <th className="px-6 py-3 text-xs font-bold text-gray-500 uppercase tracking-wider">Horário</th>
                    <th className="px-6 py-3 text-xs font-bold text-gray-500 uppercase tracking-wider">Colaborador</th>
                    <th className="px-6 py-3 text-xs font-bold text-gray-500 uppercase tracking-wider">Status</th>
                    <th className="px-6 py-3 text-xs font-bold text-gray-500 uppercase tracking-wider">Ações</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-gray-100">
                  {appointments.length > 0 ? appointments.map((app) => (
                    <tr key={app.id} className="hover:bg-gray-50 transition-colors">
                      <td className="px-6 py-4 font-medium text-gray-900">{app.time_slot}</td>
                      <td className="px-6 py-4 text-gray-700">{app.user_name}</td>
                      <td className="px-6 py-4">
                        <span className={cn("px-2 py-1 rounded-full text-[10px] font-bold uppercase border", STATUS_LABELS[app.status].color)}>
                          {STATUS_LABELS[app.status].label}
                        </span>
                      </td>
                      <td className="px-6 py-4">
                        <div className="flex items-center gap-2">
                          {app.status === 'pending' && (
                            <>
                              <button 
                                onClick={() => handleStatusChange(app.id, 'approved')}
                                className="p-1.5 text-green-600 hover:bg-green-50 rounded-lg transition-colors"
                                title="Aprovar"
                              >
                                <Check className="w-4 h-4" />
                              </button>
                              <button 
                                onClick={() => handleStatusChange(app.id, 'rejected')}
                                className="p-1.5 text-red-600 hover:bg-red-50 rounded-lg transition-colors"
                                title="Rejeitar"
                              >
                                <X className="w-4 h-4" />
                              </button>
                            </>
                          )}
                          {app.status === 'approved' && (
                            <>
                              <button 
                                onClick={() => handleStatusChange(app.id, 'completed')}
                                className="p-1.5 text-blue-600 hover:bg-blue-50 rounded-lg transition-colors"
                                title="Concluir"
                              >
                                <CheckCircle className="w-4 h-4" />
                              </button>
                              <button 
                                onClick={() => handleStatusChange(app.id, 'no-show')}
                                className="p-1.5 text-orange-600 hover:bg-orange-50 rounded-lg transition-colors"
                                title="Falta"
                              >
                                <XCircle className="w-4 h-4" />
                              </button>
                            </>
                          )}
                        </div>
                      </td>
                    </tr>
                  )) : (
                    <tr>
                      <td colSpan={4} className="px-6 py-12 text-center text-gray-400">Nenhum agendamento para este dia.</td>
                    </tr>
                  )}
                </tbody>
              </table>
            </div>
          </div>

          <div className="space-y-6">
            <div className="bg-white rounded-2xl border border-gray-200 p-6 shadow-sm">
              <h3 className="font-bold text-gray-900 mb-4 flex items-center gap-2">
                <ListOrdered className="w-5 h-5 text-blue-600" />
                Fila de Espera ({waitlist.length})
              </h3>
              <div className="space-y-3">
                {waitlist.map((entry, idx) => (
                  <div key={entry.id} className="flex items-center justify-between p-3 bg-gray-50 rounded-xl border border-gray-100">
                    <div className="flex flex-col">
                      <span className="text-sm font-medium text-gray-700">{entry.user_name}</span>
                      <span className="text-[10px] text-gray-400">Score: {entry.priority_score.toFixed(0)}</span>
                    </div>
                    <button 
                      onClick={async () => {
                        // Logic to fill a spot from waitlist
                        const availableSlot = TIME_SLOTS.find(slot => !appointments.some(a => a.time_slot === slot && ['approved', 'pending'].includes(a.status)));
                        if (availableSlot) {
                          await setDoc(doc(collection(db, 'appointments')), {
                            user_id: entry.user_id,
                            user_name: entry.user_name,
                            date: selectedDate,
                            time_slot: availableSlot,
                            status: 'approved',
                            created_at: new Date().toISOString(),
                            approved_by: profile.id
                          });
                          await deleteDoc(doc(db, 'waitlist', entry.id));
                        } else {
                          alert('Não há horários disponíveis para preenchimento automático.');
                        }
                      }}
                      className="p-2 text-blue-600 hover:bg-blue-50 rounded-lg transition-colors"
                      title="Preencher vaga"
                    >
                      <CheckCircle className="w-4 h-4" />
                    </button>
                  </div>
                ))}
              </div>
            </div>
          </div>
        </div>
      )}

      {activeTab === 'users' && (
        <div className="bg-white rounded-2xl border border-gray-200 overflow-hidden shadow-sm">
          <table className="w-full text-left border-collapse">
            <thead className="bg-gray-50 border-b border-gray-200">
              <tr>
                <th className="px-6 py-3 text-xs font-bold text-gray-500 uppercase tracking-wider">Nome</th>
                <th className="px-6 py-3 text-xs font-bold text-gray-500 uppercase tracking-wider">Score</th>
                <th className="px-6 py-3 text-xs font-bold text-gray-500 uppercase tracking-wider">Atendimentos</th>
                <th className="px-6 py-3 text-xs font-bold text-gray-500 uppercase tracking-wider">Faltas/Canc. Tardios</th>
                <th className="px-6 py-3 text-xs font-bold text-gray-500 uppercase tracking-wider">Último Atendimento</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-gray-100">
              {users.map((u) => (
                <tr key={u.id} className="hover:bg-gray-50 transition-colors">
                  <td className="px-6 py-4">
                    <div className="flex flex-col">
                      <span className="font-medium text-gray-900">{u.name}</span>
                      <span className="text-xs text-gray-500">{u.email}</span>
                    </div>
                  </td>
                  <td className="px-6 py-4">
                    <span className="font-bold text-blue-600">{u.priority_score.toFixed(0)}</span>
                  </td>
                  <td className="px-6 py-4 text-gray-700">{u.total_appointments}</td>
                  <td className="px-6 py-4 text-gray-700">
                    <span className="text-orange-600">{u.no_shows}</span> / <span className="text-red-600">{u.late_cancellations}</span>
                  </td>
                  <td className="px-6 py-4 text-sm text-gray-500">
                    {u.last_appointment_date ? format(parseISO(u.last_appointment_date), 'dd/MM/yyyy') : '-'}
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}

      {activeTab === 'import' && (
        <div className="max-w-2xl mx-auto bg-white rounded-2xl border border-gray-200 p-8 shadow-sm text-center">
          <div className="w-16 h-16 bg-blue-50 rounded-full flex items-center justify-center mx-auto mb-6">
            <FileUp className="w-8 h-8 text-blue-600" />
          </div>
          <h3 className="text-xl font-bold text-gray-900 mb-2">Importar Histórico (Excel)</h3>
          <p className="text-gray-500 mb-8">
            Selecione uma planilha com as colunas: Dia, Mês/Ano, Semana, Colaborador.
          </p>
          <label className="inline-flex items-center justify-center gap-2 bg-blue-600 text-white px-6 py-3 rounded-xl font-medium hover:bg-blue-700 transition-colors cursor-pointer shadow-sm">
            <FileUp className="w-5 h-5" />
            Selecionar Arquivo
            <input type="file" accept=".xlsx, .xls" onChange={handleImport} className="hidden" />
          </label>
        </div>
      )}

      {activeTab === 'settings' && (
        <div className="max-w-2xl mx-auto bg-white rounded-2xl border border-gray-200 p-8 shadow-sm">
          <h3 className="text-xl font-bold text-gray-900 mb-6">Configurações do Sistema</h3>
          <div className="space-y-6">
            <div>
              <label className="block text-sm font-medium text-gray-700 mb-2">Dia de Atendimento</label>
              <select 
                value={localSettings.service_day}
                onChange={(e) => setLocalSettings({ ...localSettings, service_day: parseInt(e.target.value) })}
                className="w-full border border-gray-200 rounded-xl p-3"
              >
                <option value="1">Segunda-feira</option>
                <option value="2">Terça-feira</option>
                <option value="3">Quarta-feira</option>
                <option value="4">Quinta-feira</option>
                <option value="5">Sexta-feira</option>
              </select>
            </div>
            <div>
              <label className="block text-sm font-medium text-gray-700 mb-2">Horário de Liberação da Agenda</label>
              <input 
                type="time" 
                value={localSettings.release_time}
                onChange={(e) => setLocalSettings({ ...localSettings, release_time: e.target.value })}
                className="w-full border border-gray-200 rounded-xl p-3" 
              />
            </div>
            <div className="flex items-center gap-3 p-4 bg-gray-50 rounded-xl border border-gray-100">
              <input 
                type="checkbox" 
                id="open_day_before"
                checked={localSettings.open_day_before}
                onChange={(e) => setLocalSettings({ ...localSettings, open_day_before: e.target.checked })}
                className="w-5 h-5 text-blue-600 rounded border-gray-300 focus:ring-blue-500"
              />
              <label htmlFor="open_day_before" className="text-sm font-medium text-gray-700 cursor-pointer">
                Abrir agenda um dia antes do atendimento
              </label>
            </div>
            <button 
              onClick={handleSaveSettings}
              className="w-full bg-blue-600 text-white py-3 rounded-xl font-medium hover:bg-blue-700 transition-colors"
            >
              Salvar Configurações
            </button>
          </div>
        </div>
      )}
    </motion.div>
  );
}

function UserHistory({ profile }: { profile: UserProfile }) {
  const [appointments, setAppointments] = useState<Appointment[]>([]);

  useEffect(() => {
    const q = query(collection(db, 'appointments'), where('user_id', '==', profile.id), orderBy('created_at', 'desc'));
    const unsubscribe = onSnapshot(q, (snapshot) => {
      setAppointments(snapshot.docs.map(doc => ({ id: doc.id, ...doc.data() } as Appointment)));
    }, (error) => {
      handleFirestoreError(error, OperationType.LIST, 'appointments');
    });
    return () => unsubscribe();
  }, [profile.id]);

  return (
    <motion.div 
      initial={{ opacity: 0 }}
      animate={{ opacity: 1 }}
      exit={{ opacity: 0 }}
      className="space-y-6"
    >
      <div className="bg-white rounded-2xl border border-gray-200 p-6 shadow-sm flex flex-col sm:flex-row items-center gap-6">
        <div className="w-20 h-20 bg-blue-100 rounded-full flex items-center justify-center">
          <User className="w-10 h-10 text-blue-600" />
        </div>
        <div className="flex-1 text-center sm:text-left">
          <h2 className="text-2xl font-bold text-gray-900">{profile.name}</h2>
          <p className="text-gray-500">{profile.email}</p>
          <div className="flex flex-wrap justify-center sm:justify-start gap-4 mt-4">
            <div className="bg-gray-50 px-3 py-1 rounded-lg border border-gray-100">
              <span className="text-xs text-gray-500 block">Atendimentos</span>
              <span className="font-bold text-gray-900">{profile.total_appointments}</span>
            </div>
            <div className="bg-gray-50 px-3 py-1 rounded-lg border border-gray-100">
              <span className="text-xs text-gray-500 block">Faltas</span>
              <span className="font-bold text-orange-600">{profile.no_shows}</span>
            </div>
            <div className="bg-gray-50 px-3 py-1 rounded-lg border border-gray-100">
              <span className="text-xs text-gray-500 block">Score</span>
              <span className="font-bold text-blue-600">{profile.priority_score.toFixed(0)}</span>
            </div>
          </div>
        </div>
      </div>

      <div className="bg-white rounded-2xl border border-gray-200 overflow-hidden shadow-sm">
        <div className="px-6 py-4 border-b border-gray-100 flex items-center gap-2">
          <History className="w-5 h-5 text-gray-400" />
          <h3 className="font-bold text-gray-900">Histórico de Agendamentos</h3>
        </div>
        <div className="divide-y divide-gray-100">
          {appointments.length > 0 ? appointments.map((app) => (
            <div key={app.id} className="px-6 py-4 flex items-center justify-between hover:bg-gray-50 transition-colors">
              <div className="flex items-center gap-4">
                <div className="w-10 h-10 bg-gray-100 rounded-lg flex flex-col items-center justify-center text-gray-500">
                  <span className="text-[10px] font-bold uppercase">{format(parseISO(app.date), 'MMM')}</span>
                  <span className="text-sm font-bold leading-none">{format(parseISO(app.date), 'dd')}</span>
                </div>
                <div>
                  <div className="font-medium text-gray-900">{app.time_slot}</div>
                  <div className="text-xs text-gray-500">{format(parseISO(app.date), 'dd/MM/yyyy')}</div>
                </div>
              </div>
              <span className={cn("px-2 py-1 rounded-full text-[10px] font-bold uppercase border", STATUS_LABELS[app.status].color)}>
                {STATUS_LABELS[app.status].label}
              </span>
            </div>
          )) : (
            <div className="px-6 py-12 text-center text-gray-400">Você ainda não possui agendamentos.</div>
          )}
        </div>
      </div>
    </motion.div>
  );
}

// --- Logic Helpers ---

function calculateScore(profile: Partial<UserProfile>): number {
  const baseScore = 100;
  const lastAppDate = profile.last_appointment_date ? parseISO(profile.last_appointment_date) : new Date(2000, 0, 1);
  const daysSinceLast = differenceInDays(new Date(), lastAppDate);
  
  // Rules:
  // + 10 points per day since last appointment
  // + 100 / (total_appointments + 1) to favor those with fewer appointments
  // - 50 points per no-show
  // - 20 points per late cancellation
  // - 5 points per rejection
  
  let score = baseScore;
  score += Math.min(daysSinceLast, 60) * 5; // Cap at 60 days to avoid extreme scores
  score += (100 / ((profile.total_appointments || 0) + 1));
  score -= (profile.no_shows || 0) * 50;
  score -= (profile.late_cancellations || 0) * 20;
  score -= (profile.rejections || 0) * 5;
  
  return Math.max(0, score);
}
