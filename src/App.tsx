import React, { useState, useEffect, useRef, Component } from 'react';
import { 
  User as UserIcon, 
  Calendar as CalendarIcon, 
  Table as TableIcon, 
  Camera as CameraIcon, 
  LogOut, 
  CheckCircle, 
  XCircle, 
  AlertCircle, 
  Download, 
  Plus, 
  ArrowRight, 
  Clock,
  UserPlus,
  Shield,
  FileText,
  ChevronLeft,
  ChevronRight,
  LogIn,
  Edit2,
  Trash2,
  Upload,
  Image as ImageIcon,
  Menu,
  X
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';
import { format, startOfMonth, endOfMonth, eachDayOfInterval, isSameDay, addMonths, subMonths, isToday, parseISO } from 'date-fns';
import * as XLSX from 'xlsx';
import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';
import { clsx, type ClassValue } from 'clsx';
import { twMerge } from 'tailwind-merge';

import { 
  auth, 
  db, 
  googleProvider, 
  OperationType, 
  handleFirestoreError,
  signInWithEmailAndPassword,
  createSecondaryUser
} from './firebase';
import { 
  signInWithPopup, 
  onAuthStateChanged, 
  signOut 
} from 'firebase/auth';
import { 
  collection, 
  doc, 
  getDoc, 
  getDocs, 
  setDoc, 
  addDoc, 
  updateDoc, 
  query, 
  where, 
  onSnapshot,
  deleteDoc
} from 'firebase/firestore';

function cn(...inputs: ClassValue[]) {
  return twMerge(clsx(inputs));
}

// Types
import { User, AttendanceRecord, GapReason, verifyFace } from './types';

// Components
const compressImage = (base64Str: string, maxWidth = 400, maxHeight = 400): Promise<string> => {
  return new Promise((resolve) => {
    const img = new Image();
    img.src = base64Str;
    img.onload = () => {
      const canvas = document.createElement('canvas');
      let width = img.width;
      let height = img.height;

      if (width > height) {
        if (width > maxWidth) {
          height *= maxWidth / width;
          width = maxWidth;
        }
      } else {
        if (height > maxHeight) {
          width *= maxHeight / height;
          height = maxHeight;
        }
      }

      canvas.width = width;
      canvas.height = height;
      const ctx = canvas.getContext('2d');
      ctx?.drawImage(img, 0, 0, width, height);
      resolve(canvas.toDataURL('image/jpeg', 0.7));
    };
  });
};

const Camera = ({ onCapture, registeredFace }: { onCapture: (img: string) => void, registeredFace?: string }) => {
  const videoRef = useRef<HTMLVideoElement>(null);
  const canvasRef = useRef<HTMLCanvasElement>(null);
  const [isVerifying, setIsVerifying] = useState(false);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    navigator.mediaDevices.getUserMedia({ video: true })
      .then(stream => {
        if (videoRef.current) videoRef.current.srcObject = stream;
      })
      .catch(() => setError("Camera access denied"));
  }, []);

  const capture = async () => {
    if (videoRef.current && canvasRef.current) {
      const video = videoRef.current;
      const canvas = canvasRef.current;
      const context = canvas.getContext('2d');
      if (context) {
        // Calculate crop to maintain 1:1 aspect ratio
        const videoWidth = video.videoWidth;
        const videoHeight = video.videoHeight;
        const size = Math.min(videoWidth, videoHeight);
        const startX = (videoWidth - size) / 2;
        const startY = (videoHeight - size) / 2;

        // Set canvas to square
        canvas.width = 500;
        canvas.height = 500;

        context.drawImage(video, startX, startY, size, size, 0, 0, 500, 500);

        // Add timestamp overlay
        const timestamp = format(new Date(), 'MMM dd, yyyy hh:mm:ss a');
        context.font = 'bold 16px Inter, sans-serif';
        context.fillStyle = 'rgba(0, 0, 0, 0.5)';
        const textWidth = context.measureText(timestamp).width;
        context.fillRect(10, 500 - 35, textWidth + 20, 25);
        context.fillStyle = 'white';
        context.fillText(timestamp, 20, 500 - 17);

        const rawImg = canvas.toDataURL('image/jpeg', 0.7);
        const img = await compressImage(rawImg);
        
        if (registeredFace) {
          setIsVerifying(true);
          const isMatch = await verifyFace(registeredFace, img);
          setIsVerifying(false);
          if (isMatch) {
            onCapture(img);
          } else {
            setError("Face verification failed. Please try again.");
          }
        } else {
          onCapture(img);
        }
      }
    }
  };

  return (
    <div className="flex flex-col items-center gap-4 bg-white p-4 sm:p-6 rounded-2xl shadow-xl border border-zinc-200">
      <div className="relative w-full max-w-[320px] aspect-square bg-zinc-100 rounded-xl overflow-hidden border-2 border-zinc-300 shadow-inner">
        <video ref={videoRef} autoPlay playsInline className="w-full h-full object-cover" />
        <canvas ref={canvasRef} className="hidden" />
        {isVerifying && (
          <div className="absolute inset-0 bg-black/50 flex items-center justify-center text-white">
            <div className="flex flex-col items-center gap-2">
              <div className="w-8 h-8 border-4 border-white border-t-transparent rounded-full animate-spin" />
              <p className="text-sm font-medium">Verifying Face...</p>
            </div>
          </div>
        )}
      </div>
      {error && <p className="text-red-500 text-sm font-medium">{error}</p>}
      <button 
        onClick={capture}
        disabled={isVerifying}
        className="flex items-center gap-2 bg-zinc-900 text-white px-6 py-3 rounded-full hover:bg-zinc-800 transition-all disabled:opacity-50"
      >
        <CameraIcon size={20} />
        Capture & Verify
      </button>
    </div>
  );
};

const Calendar = ({ records, userId }: { records: AttendanceRecord[], userId: string }) => {
  const [currentMonth, setCurrentMonth] = useState(new Date());
  const monthStart = startOfMonth(currentMonth);
  const monthEnd = endOfMonth(currentMonth);
  const days = eachDayOfInterval({ start: monthStart, end: monthEnd });

  const userRecords = records.filter(r => r.userId === userId);

  return (
    <div className="bg-white p-4 sm:p-6 rounded-2xl shadow-sm border border-zinc-200">
      <div className="flex items-center justify-between mb-6">
        <h3 className="text-lg sm:text-xl font-semibold text-zinc-900">{format(currentMonth, 'MMMM yyyy')}</h3>
        <div className="flex gap-2">
          <button onClick={() => setCurrentMonth(subMonths(currentMonth, 1))} className="p-2 hover:bg-zinc-100 rounded-full transition-colors">
            <ChevronLeft size={18} className="sm:w-5 sm:h-5" />
          </button>
          <button onClick={() => setCurrentMonth(addMonths(currentMonth, 1))} className="p-2 hover:bg-zinc-100 rounded-full transition-colors">
            <ChevronRight size={18} className="sm:w-5 sm:h-5" />
          </button>
        </div>
      </div>
      <div className="grid grid-cols-7 gap-1 sm:gap-2">
        {['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'].map(day => (
          <div key={day} className="text-center text-[10px] sm:text-xs font-bold text-zinc-400 uppercase tracking-wider py-2">{day}</div>
        ))}
        {Array.from({ length: monthStart.getDay() }).map((_, i) => (
          <div key={`empty-${i}`} className="aspect-square" />
        ))}
        {days.map(day => {
          const record = userRecords.find(r => isSameDay(parseISO(r.date), day));
          return (
            <div 
              key={day.toISOString()} 
              className={cn(
                "aspect-square rounded-lg sm:rounded-xl border border-zinc-100 p-1 sm:p-2 flex flex-col items-center justify-center gap-0.5 sm:gap-1 transition-all",
                isToday(day) ? "bg-zinc-50 border-zinc-300" : "hover:bg-zinc-50"
              )}
            >
              <span className={cn("text-xs sm:text-sm font-medium", isToday(day) ? "text-zinc-900" : "text-zinc-500")}>{format(day, 'd')}</span>
              {record && (
                <div className="flex gap-0.5 sm:gap-1">
                  {record.timeIn && <div className="w-1 sm:w-1.5 h-1 sm:h-1.5 rounded-full bg-emerald-500" title="Timed In" />}
                  {record.timeOut && <div className="w-1 sm:w-1.5 h-1 sm:h-1.5 rounded-full bg-blue-500" title="Timed Out" />}
                  {record.status === 'missed' && <div className="w-1 sm:w-1.5 h-1 sm:h-1.5 rounded-full bg-amber-500" title="Missed Record" />}
                </div>
              )}
            </div>
          );
        })}
      </div>
    </div>
  );
};

const AdminDashboard = ({ user, onLogout }: { user: User, onLogout: () => void }) => {
  const [view, setView] = useState<'overview' | 'employees' | 'admins' | 'attendance' | 'gaps'>('overview');
  const [users, setUsers] = useState<User[]>([]);
  const [records, setRecords] = useState<AttendanceRecord[]>([]);
  const [gaps, setGaps] = useState<GapReason[]>([]);
  const [showAddModal, setShowAddModal] = useState(false);
  const [editingUser, setEditingUser] = useState<User | null>(null);
  const [modalRole, setModalRole] = useState<'employee' | 'admin'>('employee');
  const [selectedDate, setSelectedDate] = useState(new Date());
  const [statusFilter, setStatusFilter] = useState<'all' | 'present' | 'absent' | 'timed-out'>('all');
  const [isMobileMenuOpen, setIsMobileMenuOpen] = useState(false);
  const [alertMessage, setAlertMessage] = useState<string | null>(null);
  const [deletingUserId, setDeletingUserId] = useState<string | null>(null);
  const [newUser, setNewUser] = useState({ 
    name: '', 
    email: '', 
    username: '',
    password: '', 
    authMethod: 'email' as 'email' | 'google',
    faceImage: '', 
    employeeType: 'regular' as 'regular' | 'intern',
    requiredHours: 0
  });
  const [creating, setCreating] = useState(false);
  const [externalRecords, setExternalRecords] = useState<any[]>([]);
  const [selectedPhoto, setSelectedPhoto] = useState<string | null>(null);
  const [isFetchingExternal, setIsFetchingExternal] = useState(false);

  const fetchExternal = async () => {
    setIsFetchingExternal(true);
    try {
      const res = await fetch("https://insgi-be.vercel.app/api/attendance");
      const data = await res.json();
      if (Array.isArray(data)) {
        setExternalRecords(data);
      }
    } catch (err) {
      console.error("Error fetching external records:", err);
    } finally {
      setIsFetchingExternal(false);
    }
  };

  useEffect(() => {
    if (view === 'attendance' || view === 'overview') {
      fetchExternal();
    }
  }, [view]);

  useEffect(() => {
    setIsMobileMenuOpen(false);
  }, [view]);

  useEffect(() => {
    const unsubUsers = onSnapshot(collection(db, 'users'), (snapshot) => {
      setUsers(snapshot.docs.map(doc => ({ id: doc.id, ...doc.data() } as User)));
    }, (err) => handleFirestoreError(err, OperationType.LIST, 'users'));

    const unsubRecords = onSnapshot(collection(db, 'records'), (snapshot) => {
      setRecords(snapshot.docs.map(doc => ({ id: doc.id, ...doc.data() } as AttendanceRecord)));
    }, (err) => handleFirestoreError(err, OperationType.LIST, 'records'));

    const unsubGaps = onSnapshot(collection(db, 'gaps'), (snapshot) => {
      setGaps(snapshot.docs.map(doc => ({ id: doc.id, ...doc.data() } as GapReason)));
    }, (err) => handleFirestoreError(err, OperationType.LIST, 'gaps'));

    return () => { unsubUsers(); unsubRecords(); unsubGaps(); };
  }, []);

  const resetForm = () => {
    setNewUser({ 
      name: '', 
      email: '', 
      username: '',
      password: '', 
      authMethod: 'email',
      faceImage: '', 
      employeeType: 'regular',
      requiredHours: 0
    });
    setEditingUser(null);
    setShowAddModal(false);
  };

  const handleCreateUser = async () => {
    if (!newUser.name || !newUser.email || !newUser.username) {
      setAlertMessage("Please fill in all required fields (Name, Email, Username).");
      return;
    }

    const userId = editingUser?.id || '';
    const isNewAuthUserNeeded = !editingUser || (editingUser.authMethod === 'google' && newUser.authMethod === 'email' && userId.startsWith('pending_'));

    if (newUser.authMethod === 'email' && !newUser.password && isNewAuthUserNeeded) {
      setAlertMessage("Password is required for new email accounts.");
      return;
    }
    if (modalRole === 'employee' && !newUser.faceImage && !editingUser) {
      setAlertMessage("Face image is required for new employees.");
      return;
    }
    
    setCreating(true);
    try {
      // Check if username is unique (only if creating or username changed)
      if (!editingUser || editingUser.username !== newUser.username) {
        const q = query(collection(db, 'users'), where('username', '==', newUser.username));
        const snap = await getDocs(q);
        if (!snap.empty) {
          throw new Error('Username already exists. Please choose another one.');
        }
      }

      // Check if email is unique (only if creating or email changed)
      if (!editingUser || editingUser.email !== newUser.email) {
        const q = query(collection(db, 'users'), where('email', '==', newUser.email));
        const snap = await getDocs(q);
        if (!snap.empty) {
          throw new Error('Email already exists in the system.');
        }
      }

      let finalUserId = userId;
      
      if (isNewAuthUserNeeded) {
        if (newUser.authMethod === 'email') {
          const authUser = await createSecondaryUser(newUser.email, newUser.password);
          const oldUserId = finalUserId;
          finalUserId = authUser.uid;
          
          // If we migrated from a pending user, delete the old document
          if (oldUserId && oldUserId.startsWith('pending_') && oldUserId !== finalUserId) {
            await deleteDoc(doc(db, 'users', oldUserId));
          }
        } else if (!editingUser) {
          finalUserId = `pending_${newUser.email}`;
        }
      }
      
      const userDoc: any = {
        name: newUser.name,
        email: newUser.email,
        username: newUser.username,
        role: modalRole,
        authMethod: newUser.authMethod,
        updatedAt: new Date().toISOString()
      };

      if (!editingUser) {
        userDoc.createdAt = new Date().toISOString();
      }

      if (modalRole === 'employee') {
        userDoc.faceImage = newUser.faceImage;
        userDoc.employeeType = newUser.employeeType;
        if (newUser.employeeType === 'intern') {
          userDoc.requiredHours = newUser.requiredHours;
        } else {
          userDoc.requiredHours = 0;
        }
      }

      await setDoc(doc(db, 'users', finalUserId), userDoc, { merge: true });

      // Maintain usernames collection for login lookup
      if (!editingUser || editingUser.username !== newUser.username) {
        // Delete old username if it changed
        if (editingUser && editingUser.username) {
          await deleteDoc(doc(db, 'usernames', editingUser.username));
        }
        // Set new username mapping
        await setDoc(doc(db, 'usernames', newUser.username), { email: newUser.email, userId: finalUserId });
      } else if (editingUser && editingUser.email !== newUser.email) {
        // Update email in username mapping if email changed but username stayed same
        await setDoc(doc(db, 'usernames', newUser.username), { email: newUser.email, userId: finalUserId });
      }

      resetForm();
    } catch (err: any) {
      console.error(err);
      let message = 'Failed to save account';
      if (err instanceof Error) {
        message = err.message;
        if (err.message.includes('auth/email-already-in-use')) {
          message = 'This email is already registered in the system. Please use a different email or check if the account already exists.';
        }
      }
      setAlertMessage(message);
    } finally {
      setCreating(false);
    }
  };

  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) {
      if (file.size > 5 * 1024 * 1024) {
        setAlertMessage("File size too large. Please upload an image smaller than 5MB.");
        return;
      }
      const reader = new FileReader();
      reader.onloadend = async () => {
        const compressed = await compressImage(reader.result as string);
        setNewUser({ ...newUser, faceImage: compressed });
      };
      reader.readAsDataURL(file);
    }
  };

  const handleEditUser = (user: User) => {
    setEditingUser(user);
    setModalRole(user.role);
    setNewUser({
      name: user.name,
      email: user.email,
      username: user.username || '',
      password: '',
      authMethod: user.authMethod || 'email',
      faceImage: user.faceImage || '',
      employeeType: user.employeeType || 'regular',
      requiredHours: user.requiredHours || 0
    });
    setShowAddModal(true);
  };

  const handleDeleteUser = async (userId: string) => {
    try {
      const userToDelete = users.find(u => u.id === userId);
      if (userToDelete?.username) {
        await deleteDoc(doc(db, 'usernames', userToDelete.username));
      }
      await deleteDoc(doc(db, 'users', userId));
      setDeletingUserId(null);
    } catch (err) {
      handleFirestoreError(err, OperationType.DELETE, `users/${userId}`);
    }
  };

  const employees = users.filter(u => u.role === 'employee');
  const admins = users.filter(u => u.role === 'admin');

  const navItems = [
    { id: 'overview', label: 'Overview', icon: TableIcon },
    { id: 'employees', label: 'Employees', icon: UserPlus },
    { id: 'admins', label: 'Admins', icon: Shield },
    { id: 'attendance', label: 'Attendance', icon: CalendarIcon },
    { id: 'gaps', label: 'Gap Requests', icon: AlertCircle },
  ];

  const handleExport = async () => {
    try {
      const workbook = new ExcelJS.Workbook();
      const worksheet = workbook.addWorksheet('Attendance');

      worksheet.columns = [
        { header: 'Employee', key: 'name', width: 25 },
        { header: 'Date', key: 'date', width: 15 },
        { header: 'Time In', key: 'timeIn', width: 15 },
        { header: 'Time Out', key: 'timeOut', width: 15 },
        { header: 'Status', key: 'status', width: 15 },
        { header: 'Time In Photo', key: 'inPhoto', width: 15 },
        { header: 'Time Out Photo', key: 'outPhoto', width: 15 }
      ];

      worksheet.getRow(1).font = { bold: true };
      worksheet.getRow(1).height = 20;

      for (const [index, r] of records.entries()) {
        const emp = employees.find(e => e.id === r.userId);
        const extRecs = externalRecords.filter(er => 
          String(er.id) === String(r.userId) && 
          String(er.date) === String(r.date)
        );
        
        const inPhoto = extRecs.find(er => er.time_in_photo || er.timeInPhoto || (er.photo && er.time_in))?.time_in_photo || 
                        extRecs.find(er => er.photo && er.time_in)?.photo;
                        
        const outPhoto = extRecs.find(er => er.time_out_photo || er.timeOutPhoto || (er.photo && er.time_out))?.time_out_photo ||
                         extRecs.find(er => er.photo && er.time_out)?.photo;

        const row = worksheet.addRow({
          name: emp?.name || 'Unknown',
          date: r.date,
          timeIn: r.timeIn ? format(parseISO(r.timeIn), 'hh:mm a') : '-',
          timeOut: r.timeOut ? format(parseISO(r.timeOut), 'hh:mm a') : '-',
          status: r.status
        });
        
        row.height = 80; // Set row height for images
        row.alignment = { vertical: 'middle', horizontal: 'left' };

        if (inPhoto && inPhoto.startsWith('data:image')) {
          try {
            const base64 = inPhoto.split(',')[1];
            const imageId = workbook.addImage({
              base64: base64,
              extension: 'jpeg',
            });
            worksheet.addImage(imageId, {
              tl: { col: 5, row: index + 1 },
              ext: { width: 100, height: 100 }
            });
          } catch (e) {
            console.error("Error adding inPhoto to excel", e);
          }
        }

        if (outPhoto && outPhoto.startsWith('data:image')) {
          try {
            const base64 = outPhoto.split(',')[1];
            const imageId = workbook.addImage({
              base64: base64,
              extension: 'jpeg',
            });
            worksheet.addImage(imageId, {
              tl: { col: 6, row: index + 1 },
              ext: { width: 100, height: 100 }
            });
          } catch (e) {
            console.error("Error adding outPhoto to excel", e);
          }
        }
      }

      const buffer = await workbook.xlsx.writeBuffer();
      saveAs(new Blob([buffer]), `Attendance_Report_${format(new Date(), 'yyyy-MM-dd')}.xlsx`);
    } catch (error) {
      console.error("Export failed:", error);
      setAlertMessage("Failed to export Excel file. Please try again.");
    }
  };

  const handleGapAction = async (gapId: string, status: 'approved' | 'denied') => {
    try {
      await updateDoc(doc(db, 'gaps', gapId), { status });
    } catch (err) {
      handleFirestoreError(err, OperationType.UPDATE, `gaps/${gapId}`);
    }
  };

  return (
    <div className="min-h-screen bg-zinc-50 flex flex-col lg:flex-row">
      {/* Mobile Header */}
      <div className="lg:hidden bg-white border-b border-zinc-200 p-4 flex items-center justify-between sticky top-0 z-40">
        <div className="flex items-center gap-3">
          <div className="w-10 h-10 bg-zinc-900 rounded-xl flex items-center justify-center text-white">
            <Shield size={24} />
          </div>
          <h1 className="font-bold text-zinc-900">Admin Panel</h1>
        </div>
        <button 
          onClick={() => setIsMobileMenuOpen(!isMobileMenuOpen)}
          className="p-2 hover:bg-zinc-100 rounded-xl transition-colors"
        >
          {isMobileMenuOpen ? <X size={24} /> : <Menu size={24} />}
        </button>
      </div>

      {/* Sidebar Overlay */}
      <AnimatePresence>
        {isMobileMenuOpen && (
          <motion.div 
            initial={{ opacity: 0 }}
            animate={{ opacity: 1 }}
            exit={{ opacity: 0 }}
            onClick={() => setIsMobileMenuOpen(false)}
            className="fixed inset-0 bg-black/50 z-40 lg:hidden backdrop-blur-sm"
          />
        )}
      </AnimatePresence>

      {/* Sidebar */}
      <div className={cn(
        "fixed inset-y-0 left-0 w-72 bg-white border-r border-zinc-200 flex flex-col p-6 z-50 lg:sticky lg:h-screen transition-transform duration-300 lg:translate-x-0",
        isMobileMenuOpen ? "translate-x-0" : "-translate-x-full"
      )}>
        <div className="hidden lg:flex items-center gap-4 mb-10 px-2">
          <div className="w-12 h-12 bg-zinc-900 rounded-2xl flex items-center justify-center text-white shadow-xl">
            <Shield size={28} />
          </div>
          <div>
            <h1 className="text-xl font-bold text-zinc-900">Admin Panel</h1>
            <p className="text-xs font-bold text-zinc-400 uppercase tracking-widest">Fastock Attendance</p>
          </div>
        </div>

        <nav className="flex flex-col gap-2">
          {navItems.map(item => (
            <button 
              key={item.id}
              onClick={() => setView(item.id as any)}
              className={cn(
                "flex items-center gap-3 px-4 py-3 rounded-xl transition-all font-medium", 
                view === item.id ? "bg-zinc-900 text-white shadow-lg" : "text-zinc-500 hover:bg-zinc-100"
              )}
            >
              <item.icon size={20} />
              {item.label}
            </button>
          ))}
        </nav>

        <div className="mt-auto pt-6 border-t border-zinc-100">
          <div className="px-4 py-3 mb-4 bg-zinc-50 rounded-xl">
            <p className="text-xs font-bold text-zinc-400 uppercase tracking-widest mb-1">Logged in as</p>
            <p className="text-sm font-bold text-zinc-900 truncate">{user.name}</p>
          </div>
          <button onClick={onLogout} className="flex items-center gap-3 px-4 py-3 rounded-xl text-red-500 hover:bg-red-50 transition-all font-medium w-full">
            <LogOut size={20} />
            Logout
          </button>
        </div>
      </div>

      {/* Main Content */}
      <div className="flex-1 p-4 md:p-10 overflow-auto">
        <div className="max-w-6xl mx-auto">
          <header className="flex flex-col sm:flex-row sm:items-center justify-between gap-6 mb-10">
            <div>
              <h2 className="text-2xl md:text-3xl font-bold text-zinc-900 capitalize">{view}</h2>
              <p className="text-sm md:text-base text-zinc-500">Manage your organization's attendance system</p>
            </div>
            <div className="flex gap-3">
              {view === 'employees' && (
                <button 
                  onClick={() => { resetForm(); setModalRole('employee'); setShowAddModal(true); }}
                  className="flex-1 sm:flex-none flex items-center justify-center gap-2 bg-zinc-900 text-white px-6 py-3 rounded-full hover:bg-zinc-800 transition-all shadow-lg font-bold"
                >
                  <Plus size={20} />
                  Add Employee
                </button>
              )}
              {view === 'admins' && (
                <button 
                  onClick={() => { resetForm(); setModalRole('admin'); setShowAddModal(true); }}
                  className="flex-1 sm:flex-none flex items-center justify-center gap-2 bg-zinc-900 text-white px-6 py-3 rounded-full hover:bg-zinc-800 transition-all shadow-lg font-bold"
                >
                  <Plus size={20} />
                  Add Admin
                </button>
              )}
              {view === 'attendance' && (
                <button 
                  onClick={handleExport}
                  className="flex-1 sm:flex-none flex items-center justify-center gap-2 bg-emerald-600 text-white px-6 py-3 rounded-full hover:bg-emerald-700 transition-all shadow-lg font-bold"
                >
                  <Download size={20} />
                  Export Excel
                </button>
              )}
            </div>
          </header>

          <AnimatePresence mode="wait">
            {view === 'overview' && (
              <motion.div 
                key="overview"
                initial={{ opacity: 0, y: 20 }}
                animate={{ opacity: 1, y: 0 }}
                exit={{ opacity: 0, y: -20 }}
                className="flex flex-col gap-8"
              >
                {/* Stats Grid */}
                <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
                  <div className="bg-white p-6 rounded-2xl shadow-sm border border-zinc-200">
                    <div className="flex items-center gap-4 mb-4">
                      <div className="w-12 h-12 bg-zinc-100 rounded-xl flex items-center justify-center text-zinc-600">
                        <UserIcon size={24} />
                      </div>
                      <div>
                        <p className="text-sm font-bold text-zinc-400 uppercase tracking-wider">Total People</p>
                        <h4 className="text-2xl font-bold text-zinc-900">{users.length}</h4>
                      </div>
                    </div>
                    <div className="flex gap-3 mt-1">
                      <span className="text-[10px] font-bold text-zinc-400 uppercase tracking-widest bg-zinc-50 px-2 py-0.5 rounded">{admins.length} Admins</span>
                      <span className="text-[10px] font-bold text-zinc-400 uppercase tracking-widest bg-zinc-50 px-2 py-0.5 rounded">{employees.length} Employees</span>
                    </div>
                  </div>

                  <div className="bg-white p-6 rounded-2xl shadow-sm border border-zinc-200">
                    <div className="flex items-center gap-4 mb-4">
                      <div className="w-12 h-12 bg-emerald-50 rounded-xl flex items-center justify-center text-emerald-600">
                        <Clock size={24} />
                      </div>
                      <div>
                        <p className="text-sm font-bold text-zinc-400 uppercase tracking-wider">Present Today</p>
                        <h4 className="text-2xl font-bold text-zinc-900">
                          {records.filter(r => r.date === format(new Date(), 'yyyy-MM-dd') && r.timeIn && !r.timeOut).length}
                        </h4>
                      </div>
                    </div>
                    <div className="flex items-center gap-2">
                      <div className="w-2 h-2 rounded-full bg-emerald-500 animate-pulse" />
                      <p className="text-xs font-medium text-zinc-500">Currently clocked in</p>
                    </div>
                  </div>

                  <div className="bg-white p-6 rounded-2xl shadow-sm border border-zinc-200">
                    <div className="flex items-center gap-4 mb-4">
                      <div className="w-12 h-12 bg-blue-50 rounded-xl flex items-center justify-center text-blue-600">
                        <LogOut size={24} />
                      </div>
                      <div>
                        <p className="text-sm font-bold text-zinc-400 uppercase tracking-wider">Timed Out Today</p>
                        <h4 className="text-2xl font-bold text-zinc-900">
                          {records.filter(r => r.date === format(new Date(), 'yyyy-MM-dd') && r.timeOut).length}
                        </h4>
                      </div>
                    </div>
                    <p className="text-xs font-medium text-zinc-500">Completed shifts today</p>
                  </div>
                </div>

                {/* Calendar & Daily List */}
                <div className="grid grid-cols-1 lg:grid-cols-3 gap-8">
                  <div className="lg:col-span-2">
                    <div className="bg-white p-4 sm:p-6 rounded-2xl shadow-sm border border-zinc-200">
                      <div className="flex items-center justify-between mb-6">
                        <h3 className="text-lg sm:text-xl font-bold text-zinc-900">Attendance Calendar</h3>
                        <div className="flex gap-2">
                          <button onClick={() => setSelectedDate(subMonths(selectedDate, 1))} className="p-2 hover:bg-zinc-100 rounded-full transition-colors">
                            <ChevronLeft size={18} className="sm:w-5 sm:h-5" />
                          </button>
                          <button onClick={() => setSelectedDate(addMonths(selectedDate, 1))} className="p-2 hover:bg-zinc-100 rounded-full transition-colors">
                            <ChevronRight size={18} className="sm:w-5 sm:h-5" />
                          </button>
                        </div>
                      </div>
                      <div className="grid grid-cols-7 gap-1 sm:gap-2">
                        {['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'].map(day => (
                          <div key={day} className="text-center text-[10px] sm:text-xs font-bold text-zinc-400 uppercase tracking-wider py-2">{day}</div>
                        ))}
                        {Array.from({ length: startOfMonth(selectedDate).getDay() }).map((_, i) => (
                          <div key={`empty-${i}`} className="aspect-square" />
                        ))}
                        {eachDayOfInterval({ start: startOfMonth(selectedDate), end: endOfMonth(selectedDate) }).map(day => {
                          const dateStr = format(day, 'yyyy-MM-dd');
                          const dayRecords = records.filter(r => r.date === dateStr);
                          const isSelected = isSameDay(day, selectedDate);
                          
                          return (
                            <button 
                              key={day.toISOString()} 
                              onClick={() => setSelectedDate(day)}
                              className={cn(
                                "aspect-square rounded-lg sm:rounded-xl border p-1 sm:p-2 flex flex-col items-center justify-center gap-0.5 sm:gap-1 transition-all relative",
                                isSelected ? "bg-zinc-900 border-zinc-900 text-white shadow-lg" : "border-zinc-100 hover:bg-zinc-50 text-zinc-500",
                                isToday(day) && !isSelected && "border-zinc-300 bg-zinc-50"
                              )}
                            >
                              <span className="text-xs sm:text-sm font-bold">{format(day, 'd')}</span>
                              <div className="flex gap-0.5">
                                {dayRecords.some(r => r.status === 'present') && (
                                  <div className={cn("w-1 h-1 rounded-full", isSelected ? "bg-emerald-400" : "bg-emerald-500")} />
                                )}
                                {dayRecords.some(r => r.status === 'missed') && (
                                  <div className={cn("w-1 h-1 rounded-full", isSelected ? "bg-amber-400" : "bg-amber-500")} />
                                )}
                              </div>
                            </button>
                          );
                        })}
                      </div>
                      
                      {/* Calendar Legend */}
                      <div className="mt-6 pt-6 border-t border-zinc-100 flex items-center gap-6">
                        <div className="flex items-center gap-2">
                          <div className="w-2 h-2 rounded-full bg-emerald-500" />
                          <span className="text-xs font-bold text-zinc-400 uppercase tracking-wider">Present</span>
                        </div>
                        <div className="flex items-center gap-2">
                          <div className="w-2 h-2 rounded-full bg-amber-500" />
                          <span className="text-xs font-bold text-zinc-400 uppercase tracking-wider">Missed/Absent</span>
                        </div>
                        <div className="flex items-center gap-2">
                          <div className="w-2 h-2 rounded-full bg-zinc-300" />
                          <span className="text-xs font-bold text-zinc-400 uppercase tracking-wider">No Record</span>
                        </div>
                      </div>
                    </div>
                  </div>
                  <div className="lg:col-span-1">
                    <div className="bg-white p-4 sm:p-6 rounded-2xl shadow-sm border border-zinc-200 h-full flex flex-col">
                      <div className="flex items-center justify-between mb-6">
                        <h3 className="text-lg font-bold text-zinc-900">
                          {format(selectedDate, 'MMM d, yyyy')}
                        </h3>
                        <select 
                          value={statusFilter}
                          onChange={(e) => setStatusFilter(e.target.value as any)}
                          className="text-xs font-bold border-none bg-zinc-100 rounded-lg px-2 py-1 focus:ring-0"
                        >
                          <option value="all">All</option>
                          <option value="present">Present</option>
                          <option value="absent">Absent</option>
                          <option value="timed-out">Timed Out</option>
                        </select>
                      </div>
                      
                      <div className="space-y-3 overflow-auto flex-1 max-h-[500px] pr-2 scrollbar-hide">
                        {employees.filter(emp => {
                          const record = records.find(r => r.userId === emp.id && r.date === format(selectedDate, 'yyyy-MM-dd'));
                          const isTodaySelected = isToday(selectedDate);
                          
                          if (statusFilter === 'all') return true;
                          if (statusFilter === 'present') return record?.timeIn && !record?.timeOut;
                          if (statusFilter === 'timed-out') return record?.timeOut;
                          if (statusFilter === 'absent') return !record || (isTodaySelected ? false : record.status === 'missed');
                          return true;
                        }).map(emp => {
                          const record = records.find(r => r.userId === emp.id && r.date === format(selectedDate, 'yyyy-MM-dd'));
                          const isTodaySelected = isToday(selectedDate);
                          
                          let statusLabel = 'Absent';
                          let statusColor = 'bg-zinc-50 text-zinc-400';
                          
                          if (record) {
                            if (record.timeOut) {
                              statusLabel = 'Timed Out';
                              statusColor = 'bg-blue-50 text-blue-600';
                            } else if (record.timeIn) {
                              statusLabel = 'Present';
                              statusColor = 'bg-emerald-50 text-emerald-600';
                            } else if (record.status === 'missed') {
                              statusLabel = 'Missed';
                              statusColor = 'bg-amber-50 text-amber-600';
                            }
                          } else if (isTodaySelected) {
                            statusLabel = 'Not Timed In';
                            statusColor = 'bg-zinc-50 text-zinc-400';
                          }

                          return (
                            <div key={emp.id} className="flex items-center justify-between p-3 rounded-xl border border-zinc-100 hover:border-zinc-200 transition-all gap-3">
                              <div className="flex items-center gap-3 min-w-0 flex-1">
                                <div className="w-8 h-8 rounded-full bg-zinc-100 flex-shrink-0 flex items-center justify-center text-zinc-500 text-xs font-bold">
                                  {emp.name.charAt(0)}
                                </div>
                                <div className="min-w-0 flex-1">
                                  <p className="text-sm font-bold text-zinc-900 truncate">{emp.name}</p>
                                  <p className="text-[10px] text-zinc-400 uppercase font-bold tracking-wider truncate">{emp.employeeType}</p>
                                </div>
                              </div>
                              <div className="text-right flex-shrink-0">
                                <span className={cn(
                                  "text-[10px] font-bold uppercase tracking-wider px-2 py-1 rounded-md whitespace-nowrap",
                                  statusColor
                                )}>
                                  {statusLabel}
                                </span>
                              </div>
                            </div>
                          );
                        })}
                        {employees.filter(emp => {
                          const record = records.find(r => r.userId === emp.id && r.date === format(selectedDate, 'yyyy-MM-dd'));
                          const isTodaySelected = isToday(selectedDate);
                          if (statusFilter === 'all') return true;
                          if (statusFilter === 'present') return record?.timeIn && !record?.timeOut;
                          if (statusFilter === 'timed-out') return record?.timeOut;
                          if (statusFilter === 'absent') return !record || (isTodaySelected ? false : record.status === 'missed');
                          return true;
                        }).length === 0 && (
                          <div className="text-center py-10">
                            <p className="text-sm text-zinc-400">No employees found for this status</p>
                          </div>
                        )}
                      </div>
                    </div>
                  </div>
                </div>
              </motion.div>
            )}

            {view === 'employees' && (
              <motion.div 
                key="employees"
                initial={{ opacity: 0, y: 20 }}
                animate={{ opacity: 1, y: 0 }}
                exit={{ opacity: 0, y: -20 }}
                className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6"
              >
                {employees.map(emp => (
                  <div key={emp.id} className="bg-white p-6 rounded-2xl shadow-sm border border-zinc-200 flex items-center justify-between group gap-4">
                    <div className="flex items-center gap-4 min-w-0 flex-1">
                      <div className="w-16 h-16 rounded-full overflow-hidden bg-zinc-100 flex-shrink-0 flex items-center justify-center">
                        {emp.faceImage ? (
                          <img src={emp.faceImage} alt={emp.name} className="w-full h-full object-cover" referrerPolicy="no-referrer" />
                        ) : (
                          <UserIcon size={32} className="text-zinc-400" />
                        )}
                      </div>
                      <div className="min-w-0 flex-1">
                        <h4 className="font-bold text-zinc-900 truncate">{emp.name}</h4>
                        <p className="text-sm text-zinc-500 truncate">{emp.email}</p>
                        <div className="flex flex-wrap gap-2 mt-1">
                          <span className="text-[10px] font-bold uppercase tracking-widest bg-zinc-100 text-zinc-500 px-2 py-0.5 rounded whitespace-nowrap">{emp.employeeType}</span>
                          {emp.username && <span className="text-[10px] font-bold uppercase tracking-widest bg-zinc-100 text-zinc-500 px-2 py-0.5 rounded whitespace-nowrap truncate max-w-[100px]">@{emp.username}</span>}
                        </div>
                      </div>
                    </div>
                    <div className="flex flex-col gap-2 lg:opacity-0 lg:group-hover:opacity-100 transition-opacity flex-shrink-0">
                      <button 
                        onClick={() => handleEditUser(emp)}
                        className="p-2 bg-zinc-100 text-zinc-600 rounded-lg hover:bg-zinc-200 transition-all"
                        title="Edit User"
                      >
                        <Edit2 size={16} />
                      </button>
                      <button 
                        onClick={() => setDeletingUserId(emp.id)}
                        className="p-2 bg-red-50 text-red-600 rounded-lg hover:bg-red-100 transition-all"
                        title="Delete User"
                      >
                        <Trash2 size={16} />
                      </button>
                    </div>
                  </div>
                ))}
              </motion.div>
            )}

            {view === 'admins' && (
              <motion.div 
                key="admins"
                initial={{ opacity: 0, y: 20 }}
                animate={{ opacity: 1, y: 0 }}
                exit={{ opacity: 0, y: -20 }}
                className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6"
              >
                {admins.map(admin => (
                  <div key={admin.id} className="bg-white p-6 rounded-2xl shadow-sm border border-zinc-200 flex items-center justify-between group gap-4">
                    <div className="flex items-center gap-4 min-w-0 flex-1">
                      <div className="w-16 h-16 rounded-full overflow-hidden bg-zinc-900 flex-shrink-0 flex items-center justify-center text-white">
                        <Shield size={32} />
                      </div>
                      <div className="min-w-0 flex-1">
                        <h4 className="font-bold text-zinc-900 truncate">{admin.name}</h4>
                        <p className="text-sm text-zinc-500 truncate">{admin.email}</p>
                        <div className="flex flex-wrap gap-2 mt-1">
                          <span className="text-[10px] font-bold uppercase tracking-widest bg-zinc-100 text-zinc-500 px-2 py-0.5 rounded whitespace-nowrap">Administrator</span>
                          {admin.username && <span className="text-[10px] font-bold uppercase tracking-widest bg-zinc-100 text-zinc-500 px-2 py-0.5 rounded whitespace-nowrap truncate max-w-[100px]">@{admin.username}</span>}
                        </div>
                      </div>
                    </div>
                    <div className="flex flex-col gap-2 lg:opacity-0 lg:group-hover:opacity-100 transition-opacity flex-shrink-0">
                      <button 
                        onClick={() => handleEditUser(admin)}
                        className="p-2 bg-zinc-100 text-zinc-600 rounded-lg hover:bg-zinc-200 transition-all"
                        title="Edit Admin"
                      >
                        <Edit2 size={16} />
                      </button>
                      {admin.role === 'admin' && admin.id !== user.id && (
                        <button 
                          onClick={() => setDeletingUserId(admin.id)}
                          className="p-2 bg-red-50 text-red-600 rounded-lg hover:bg-red-100 transition-all"
                          title="Delete Admin"
                        >
                          <Trash2 size={16} />
                        </button>
                      )}
                    </div>
                  </div>
                ))}
              </motion.div>
            )}

            {view === 'attendance' && (
              <motion.div 
                key="attendance"
                initial={{ opacity: 0, y: 20 }}
                animate={{ opacity: 1, y: 0 }}
                exit={{ opacity: 0, y: -20 }}
                className="flex flex-col gap-8"
              >
                <div className="flex justify-between items-center">
                  <h2 className="text-2xl font-bold text-zinc-900">Attendance Records</h2>
                  <button 
                    onClick={fetchExternal}
                    disabled={isFetchingExternal}
                    className="flex items-center gap-2 px-4 py-2 bg-white border border-zinc-200 rounded-xl text-sm font-medium text-zinc-600 hover:bg-zinc-50 transition-all disabled:opacity-50"
                  >
                    <Clock size={16} className={cn(isFetchingExternal && "animate-spin")} />
                    {isFetchingExternal ? 'Refreshing...' : 'Refresh Photos'}
                  </button>
                </div>
                <div className="bg-white rounded-2xl shadow-sm border border-zinc-200 overflow-x-auto scrollbar-hide">
                  <table className="w-full text-left min-w-[800px]">
                    <thead className="bg-zinc-50 border-b border-zinc-200">
                      <tr>
                        <th className="px-6 py-4 text-xs font-bold text-zinc-400 uppercase tracking-wider">Employee</th>
                        <th className="px-6 py-4 text-xs font-bold text-zinc-400 uppercase tracking-wider">Date</th>
                        <th className="px-6 py-4 text-xs font-bold text-zinc-400 uppercase tracking-wider">Time In</th>
                        <th className="px-6 py-4 text-xs font-bold text-zinc-400 uppercase tracking-wider">Time Out</th>
                        <th className="px-6 py-4 text-xs font-bold text-zinc-400 uppercase tracking-wider">Photos</th>
                        <th className="px-6 py-4 text-xs font-bold text-zinc-400 uppercase tracking-wider">Status</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-zinc-100">
                      {records.map(r => {
                        const emp = employees.find(e => e.id === r.userId);
                        // Robust matching: check both string and potentially numeric IDs, and date strings
                        const extRecs = externalRecords.filter(er => 
                          String(er.id) === String(r.userId) && 
                          String(er.date) === String(r.date)
                        );
                        
                        // Check multiple possible property names for photos
                        const inPhoto = extRecs.find(er => er.time_in_photo || er.timeInPhoto || (er.photo && er.time_in))?.time_in_photo || 
                                        extRecs.find(er => er.photo && er.time_in)?.photo;
                                        
                        const outPhoto = extRecs.find(er => er.time_out_photo || er.timeOutPhoto || (er.photo && er.time_out))?.time_out_photo ||
                                         extRecs.find(er => er.photo && er.time_out)?.photo;

                        return (
                          <tr key={r.id} className="hover:bg-zinc-50 transition-colors">
                            <td className="px-6 py-4 font-medium text-zinc-900">{emp?.name || 'Unknown'}</td>
                            <td className="px-6 py-4 text-zinc-500">{r.date}</td>
                            <td className="px-6 py-4 text-zinc-500">{r.timeIn ? format(parseISO(r.timeIn), 'hh:mm a') : '-'}</td>
                            <td className="px-6 py-4 text-zinc-500">{r.timeOut ? format(parseISO(r.timeOut), 'hh:mm a') : '-'}</td>
                            <td className="px-6 py-4">
                              <div className="flex gap-3">
                                {inPhoto ? (
                                  <div className="flex flex-col items-center gap-1">
                                    <button 
                                      onClick={() => setSelectedPhoto(inPhoto)}
                                      className="w-10 h-10 rounded-lg overflow-hidden border-2 border-emerald-100 hover:border-emerald-400 transition-all shadow-sm"
                                      title="View Time In Photo"
                                    >
                                      <img src={inPhoto} alt="In" className="w-full h-full object-cover" referrerPolicy="no-referrer" />
                                    </button>
                                    <span className="text-[9px] font-bold text-emerald-600 uppercase tracking-tighter">In</span>
                                  </div>
                                ) : (
                                  <div className="w-10 h-10 rounded-lg bg-zinc-50 border border-dashed border-zinc-200 flex items-center justify-center" title="No Time In Photo">
                                    <ImageIcon size={14} className="text-zinc-300" />
                                  </div>
                                )}
                                
                                {outPhoto ? (
                                  <div className="flex flex-col items-center gap-1">
                                    <button 
                                      onClick={() => setSelectedPhoto(outPhoto)}
                                      className="w-10 h-10 rounded-lg overflow-hidden border-2 border-blue-100 hover:border-blue-400 transition-all shadow-sm"
                                      title="View Time Out Photo"
                                    >
                                      <img src={outPhoto} alt="Out" className="w-full h-full object-cover" referrerPolicy="no-referrer" />
                                    </button>
                                    <span className="text-[9px] font-bold text-blue-600 uppercase tracking-tighter">Out</span>
                                  </div>
                                ) : (
                                  <div className="w-10 h-10 rounded-lg bg-zinc-50 border border-dashed border-zinc-200 flex items-center justify-center" title="No Time Out Photo">
                                    <ImageIcon size={14} className="text-zinc-300" />
                                  </div>
                                )}
                              </div>
                            </td>
                            <td className="px-6 py-4">
                              <span className={cn(
                                "px-3 py-1 rounded-full text-xs font-bold uppercase tracking-wider",
                                r.status === 'present' ? "bg-emerald-100 text-emerald-700" : "bg-amber-100 text-amber-700"
                              )}>
                                {r.status}
                              </span>
                            </td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>
              </motion.div>
            )}

            {view === 'gaps' && (
              <motion.div 
                key="gaps"
                initial={{ opacity: 0, y: 20 }}
                animate={{ opacity: 1, y: 0 }}
                exit={{ opacity: 0, y: -20 }}
                className="grid grid-cols-1 gap-6"
              >
                {gaps.map(gap => {
                  const emp = employees.find(e => e.id === gap.userId);
                  const record = records.find(r => r.id === gap.recordId);
                  return (
                    <div key={gap.id} className="bg-white p-6 rounded-2xl shadow-sm border border-zinc-200 flex items-center justify-between">
                      <div className="flex items-center gap-4">
                        <div className="w-12 h-12 bg-zinc-100 rounded-xl flex items-center justify-center text-zinc-400">
                          <FileText size={24} />
                        </div>
                        <div>
                          <h4 className="font-bold text-zinc-900">{emp?.name}</h4>
                          <p className="text-sm text-zinc-500">Record Date: {record?.date}</p>
                          <p className="text-sm text-zinc-600 mt-2 font-medium italic">"{gap.reason}"</p>
                        </div>
                      </div>
                      <div className="flex items-center gap-3">
                        {gap.status === 'pending' ? (
                          <>
                            <button 
                              onClick={() => handleGapAction(gap.id, 'approved')}
                              className="p-3 bg-emerald-50 text-emerald-600 rounded-xl hover:bg-emerald-100 transition-all"
                            >
                              <CheckCircle size={20} />
                            </button>
                            <button 
                              onClick={() => handleGapAction(gap.id, 'denied')}
                              className="p-3 bg-red-50 text-red-600 rounded-xl hover:bg-red-100 transition-all"
                            >
                              <XCircle size={20} />
                            </button>
                          </>
                        ) : (
                          <span className={cn(
                            "px-4 py-2 rounded-xl text-sm font-bold uppercase tracking-wider",
                            gap.status === 'approved' ? "bg-emerald-100 text-emerald-700" : "bg-red-100 text-red-700"
                          )}>
                            {gap.status}
                          </span>
                        )}
                      </div>
                    </div>
                  );
                })}
              </motion.div>
            )}
          </AnimatePresence>
        </div>
      </div>

      {/* Add User Modal */}
      {showAddModal && (
        <div className="fixed inset-0 bg-black/50 backdrop-blur-sm flex items-end sm:items-center justify-center z-50 p-0 sm:p-6">
          <motion.div 
            initial={{ y: "100%", opacity: 0 }}
            animate={{ y: 0, opacity: 1 }}
            className="bg-white rounded-t-3xl sm:rounded-3xl shadow-2xl max-w-2xl w-full overflow-hidden max-h-[90vh] flex flex-col"
          >
            <div className="p-6 sm:p-8 overflow-y-auto">
              <div className="flex items-center justify-between mb-8">
                <h3 className="text-xl sm:text-2xl font-bold text-zinc-900">
                  {editingUser ? 'Edit' : 'Add New'} {modalRole === 'admin' ? 'Admin' : 'Employee'}
                </h3>
                <button onClick={resetForm} className="p-2 hover:bg-zinc-100 rounded-full">
                  <XCircle size={24} className="text-zinc-400" />
                </button>
              </div>
              
              <div className="space-y-6">
                <div>
                  <label className="block text-xs font-bold text-zinc-400 uppercase tracking-wider mb-2">Sign-in Method</label>
                  <div className="flex bg-zinc-100 p-1 rounded-2xl">
                    <button 
                      onClick={() => setNewUser({ ...newUser, authMethod: 'email' })}
                      className={cn("flex-1 py-2 rounded-xl text-sm font-bold transition-all", newUser.authMethod === 'email' ? "bg-white text-zinc-900 shadow-sm" : "text-zinc-500")}
                    >
                      Email/Password
                    </button>
                    <button 
                      onClick={() => setNewUser({ ...newUser, authMethod: 'google' })}
                      className={cn("flex-1 py-2 rounded-xl text-sm font-bold transition-all", newUser.authMethod === 'google' ? "bg-white text-zinc-900 shadow-sm" : "text-zinc-500")}
                    >
                      Google
                    </button>
                  </div>
                </div>
                <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                  <div>
                    <label className="block text-xs font-bold text-zinc-400 uppercase tracking-wider mb-2">Username / User ID</label>
                    <input 
                      type="text" 
                      value={newUser.username}
                      onChange={e => setNewUser({ ...newUser, username: e.target.value.toLowerCase().replace(/\s+/g, '') })}
                      className="w-full px-4 py-3 rounded-xl border border-zinc-200 focus:ring-2 focus:ring-zinc-900 focus:border-transparent outline-none transition-all"
                      placeholder="johndoe123"
                    />
                  </div>
                  <div>
                    <label className="block text-xs font-bold text-zinc-400 uppercase tracking-wider mb-2">Full Name</label>
                    <input 
                      type="text" 
                      value={newUser.name}
                      onChange={e => setNewUser({ ...newUser, name: e.target.value })}
                      className="w-full px-4 py-3 rounded-xl border border-zinc-200 focus:ring-2 focus:ring-zinc-900 focus:border-transparent outline-none transition-all"
                      placeholder="John Doe"
                    />
                  </div>
                </div>
                <div>
                  <label className="block text-xs font-bold text-zinc-400 uppercase tracking-wider mb-2">Email Address</label>
                  <input 
                    type="email" 
                    value={newUser.email}
                    onChange={e => setNewUser({ ...newUser, email: e.target.value })}
                    className="w-full px-4 py-3 rounded-xl border border-zinc-200 focus:ring-2 focus:ring-zinc-900 focus:border-transparent outline-none transition-all"
                    placeholder="john@example.com"
                  />
                </div>
                {newUser.authMethod === 'email' && (
                  <div>
                    <label className="block text-xs font-bold text-zinc-400 uppercase tracking-wider mb-2">Password</label>
                    <input 
                      type="password" 
                      value={newUser.password}
                      onChange={e => setNewUser({ ...newUser, password: e.target.value })}
                      className="w-full px-4 py-3 rounded-xl border border-zinc-200 focus:ring-2 focus:ring-zinc-900 focus:border-transparent outline-none transition-all"
                      placeholder="••••••••"
                    />
                  </div>
                )}

                {modalRole === 'employee' && (
                  <>
                    <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                      <div>
                        <label className="block text-xs font-bold text-zinc-400 uppercase tracking-wider mb-2">Employee Type</label>
                        <select 
                          value={newUser.employeeType}
                          onChange={e => setNewUser({ ...newUser, employeeType: e.target.value as any })}
                          className="w-full px-4 py-3 rounded-xl border border-zinc-200 focus:ring-2 focus:ring-zinc-900 focus:border-transparent outline-none transition-all"
                        >
                          <option value="regular">Regular</option>
                          <option value="intern">Intern</option>
                        </select>
                      </div>
                      {newUser.employeeType === 'intern' && (
                        <div>
                          <label className="block text-xs font-bold text-zinc-400 uppercase tracking-wider mb-2">Required Hours</label>
                          <input 
                            type="number" 
                            value={newUser.requiredHours}
                            onChange={e => setNewUser({ ...newUser, requiredHours: parseInt(e.target.value) || 0 })}
                            className="w-full px-4 py-3 rounded-xl border border-zinc-200 focus:ring-2 focus:ring-zinc-900 focus:border-transparent outline-none transition-all"
                            placeholder="150"
                          />
                        </div>
                      )}
                    </div>
                    <div>
                      <label className="block text-xs font-bold text-zinc-400 uppercase tracking-wider mb-2">Face Registration</label>
                      {newUser.faceImage ? (
                        <div className="relative w-full aspect-video rounded-xl overflow-hidden border-2 border-zinc-200">
                          <img src={newUser.faceImage} alt="Face" className="w-full h-full object-cover" />
                          <button 
                            onClick={() => setNewUser({ ...newUser, faceImage: '' })}
                            className="absolute top-4 right-4 p-2 bg-white/90 rounded-full shadow-lg hover:bg-white transition-all"
                          >
                            <XCircle size={20} className="text-red-500" />
                          </button>
                        </div>
                      ) : (
                        <div className="space-y-4">
                          <Camera onCapture={img => setNewUser({ ...newUser, faceImage: img })} />
                          
                          <div className="relative flex items-center py-2">
                            <div className="flex-grow border-t border-zinc-200"></div>
                            <span className="flex-shrink mx-4 text-[10px] font-bold text-zinc-400 uppercase tracking-widest">OR</span>
                            <div className="flex-grow border-t border-zinc-200"></div>
                          </div>

                          <label className="flex flex-col items-center justify-center w-full h-32 border-2 border-dashed border-zinc-200 rounded-2xl hover:bg-zinc-50 transition-all cursor-pointer group">
                            <div className="flex flex-col items-center justify-center pt-5 pb-6">
                              <Upload className="w-6 h-6 text-zinc-400 group-hover:text-zinc-600 mb-2" />
                              <p className="text-xs text-zinc-500 group-hover:text-zinc-700 font-medium">Click to upload face image</p>
                              <p className="text-[10px] text-zinc-400 mt-1">PNG, JPG or JPEG (MAX. 2MB)</p>
                            </div>
                            <input 
                              type="file" 
                              className="hidden" 
                              accept="image/*"
                              onChange={handleFileUpload}
                            />
                          </label>
                        </div>
                      )}
                    </div>
                  </>
                )}
              </div>

              <div className="mt-10 flex flex-col sm:flex-row gap-4 pb-6 sm:pb-0">
                <button 
                  onClick={resetForm}
                  className="order-2 sm:order-1 flex-1 px-6 py-4 rounded-2xl border border-zinc-200 font-bold text-zinc-500 hover:bg-zinc-50 transition-all"
                >
                  Cancel
                </button>
                <button 
                  onClick={handleCreateUser}
                  disabled={creating || !newUser.name || !newUser.email || !newUser.username}
                  className="order-1 sm:order-2 flex-1 px-6 py-4 rounded-2xl bg-zinc-900 text-white font-bold hover:bg-zinc-800 transition-all shadow-lg disabled:opacity-50 flex items-center justify-center gap-2"
                >
                  {creating && <div className="w-5 h-5 border-2 border-white border-t-transparent rounded-full animate-spin" />}
                  {editingUser ? 'Save Changes' : `Create ${modalRole === 'admin' ? 'Admin' : 'Account'}`}
                </button>
              </div>
            </div>
          </motion.div>
        </div>
      )}

      {/* Alert Modal */}
      {alertMessage && (
        <div className="fixed inset-0 bg-black/50 backdrop-blur-sm flex items-center justify-center z-[100] p-6">
          <motion.div 
            initial={{ scale: 0.9, opacity: 0 }}
            animate={{ scale: 1, opacity: 1 }}
            className="bg-white rounded-3xl shadow-2xl max-w-sm w-full p-8 text-center"
          >
            <div className="w-16 h-16 bg-red-50 text-red-500 rounded-full flex items-center justify-center mx-auto mb-6">
              <AlertCircle size={32} />
            </div>
            <h3 className="text-xl font-bold text-zinc-900 mb-2">Notification</h3>
            <p className="text-zinc-500 mb-8">{alertMessage}</p>
            <button 
              onClick={() => setAlertMessage(null)}
              className="w-full bg-zinc-900 text-white py-4 rounded-2xl font-bold hover:bg-zinc-800 transition-all shadow-lg"
            >
              Close
            </button>
          </motion.div>
        </div>
      )}

      {/* Confirmation Modal */}
      {deletingUserId && (
        <div className="fixed inset-0 bg-black/50 backdrop-blur-sm flex items-center justify-center z-[100] p-6">
          <motion.div 
            initial={{ scale: 0.9, opacity: 0 }}
            animate={{ scale: 1, opacity: 1 }}
            className="bg-white rounded-3xl shadow-2xl max-w-sm w-full p-8 text-center"
          >
            <div className="w-16 h-16 bg-amber-50 text-amber-500 rounded-full flex items-center justify-center mx-auto mb-6">
              <AlertCircle size={32} />
            </div>
            <h3 className="text-xl font-bold text-zinc-900 mb-2">Confirm Delete</h3>
            <p className="text-zinc-500 mb-8">Are you sure you want to delete this user? This action cannot be undone.</p>
            <div className="flex flex-col sm:flex-row gap-3">
              <button 
                onClick={() => setDeletingUserId(null)}
                className="flex-1 px-6 py-4 rounded-2xl border border-zinc-200 font-bold text-zinc-500 hover:bg-zinc-50 transition-all"
              >
                Cancel
              </button>
              <button 
                onClick={() => handleDeleteUser(deletingUserId)}
                className="flex-1 px-6 py-4 rounded-2xl bg-red-600 text-white font-bold hover:bg-red-700 transition-all shadow-lg"
              >
                Delete
              </button>
            </div>
          </motion.div>
        </div>
      )}

      {/* Photo Viewer Modal */}
      {selectedPhoto && (
        <div className="fixed inset-0 bg-black/80 backdrop-blur-sm flex items-center justify-center z-[110] p-6" onClick={() => setSelectedPhoto(null)}>
          <motion.div 
            initial={{ scale: 0.9, opacity: 0 }}
            animate={{ scale: 1, opacity: 1 }}
            className="bg-white rounded-3xl shadow-2xl max-w-2xl w-full p-4 overflow-hidden relative"
            onClick={e => e.stopPropagation()}
          >
            <button onClick={() => setSelectedPhoto(null)} className="absolute top-6 right-6 p-2 bg-black/50 text-white rounded-full hover:bg-black/70 transition-all z-10">
              <XCircle size={24} />
            </button>
            <img src={selectedPhoto} alt="Attendance verification" className="w-full h-auto rounded-2xl" referrerPolicy="no-referrer" />
          </motion.div>
        </div>
      )}
    </div>
  );
};

const EmployeeDashboard = ({ user, onLogout }: { user: User, onLogout: () => void }) => {
  const [records, setRecords] = useState<AttendanceRecord[]>([]);
  const [gaps, setGaps] = useState<GapReason[]>([]);
  const [showCamera, setShowCamera] = useState(false);
  const [cameraMode, setCameraMode] = useState<'in' | 'out'>('in');
  const [missedRecord, setMissedRecord] = useState<AttendanceRecord | null>(null);
  const [gapReason, setGapReason] = useState('');
  const [view, setView] = useState<'calendar' | 'table'>('calendar');

  useEffect(() => {
    const unsubRecords = onSnapshot(query(collection(db, 'records'), where('userId', '==', user.id)), (snapshot) => {
      const userRecords = snapshot.docs.map(doc => ({ id: doc.id, ...doc.data() } as AttendanceRecord));
      setRecords(userRecords);

      // Check for missed time out from yesterday or before
      const today = format(new Date(), 'yyyy-MM-dd');
      const missed = userRecords.find(rec => rec.date !== today && rec.timeIn && !rec.timeOut && rec.status !== 'missed');
      if (missed) {
        setMissedRecord(missed);
      }
    }, (err) => handleFirestoreError(err, OperationType.LIST, 'records'));

    const unsubGaps = onSnapshot(query(collection(db, 'gaps'), where('userId', '==', user.id)), (snapshot) => {
      setGaps(snapshot.docs.map(doc => ({ id: doc.id, ...doc.data() } as GapReason)));
    }, (err) => handleFirestoreError(err, OperationType.LIST, 'gaps'));

    return () => { unsubRecords(); unsubGaps(); };
  }, [user.id]);

  const handleTimeAction = async (img: string) => {
    const today = format(new Date(), 'yyyy-MM-dd');
    const existingRecord = records.find(r => r.date === today);
    const now = new Date().toISOString();

    try {
      if (cameraMode === 'in') {
        if (!existingRecord) {
          await addDoc(collection(db, 'records'), {
            userId: user.id,
            date: today,
            timeIn: now,
            status: 'present'
          });

          // Send to external SQL API
          try {
            await fetch("https://insgi-be.vercel.app/api/attendance", {
              method: "POST",
              headers: { "Content-Type": "application/json" },
              body: JSON.stringify({
                id: user.id,
                time_in: now,
                time_out: null,
                time_in_photo: img,
                time_out_photo: null,
                photo: img, // Generic field for backward compatibility
                date: today
              })
            });
          } catch (apiErr) {
            console.error("External API Error (Clock In):", apiErr);
          }
        }
      } else {
        if (existingRecord && !existingRecord.timeOut) {
          await updateDoc(doc(db, 'records', existingRecord.id), {
            timeOut: now
          });

          // Send to external SQL API
          try {
            await fetch("https://insgi-be.vercel.app/api/attendance", {
              method: "POST",
              headers: { "Content-Type": "application/json" },
              body: JSON.stringify({
                id: user.id,
                time_in: existingRecord.timeIn,
                time_out: now,
                time_in_photo: null,
                time_out_photo: img,
                photo: img, // Generic field for backward compatibility
                date: today
              })
            });
          } catch (apiErr) {
            console.error("External API Error (Clock Out):", apiErr);
          }
        }
      }
    } catch (err) {
      handleFirestoreError(err, OperationType.WRITE, 'records');
    }
    setShowCamera(false);
  };

  const handleSubmitGap = async () => {
    if (!missedRecord || !gapReason) return;
    try {
      await addDoc(collection(db, 'gaps'), {
        userId: user.id,
        recordId: missedRecord.id,
        reason: gapReason,
        status: 'pending'
      });
      await updateDoc(doc(db, 'records', missedRecord.id), { status: 'missed' });
      setMissedRecord(null);
      setGapReason('');
    } catch (err) {
      handleFirestoreError(err, OperationType.WRITE, 'gaps');
    }
  };

  const todayRecord = records.find(r => r.date === format(new Date(), 'yyyy-MM-dd'));

  return (
    <div className="min-h-screen bg-zinc-50 p-4 sm:p-6 md:p-10">
      <div className="max-w-6xl mx-auto">
        <header className="flex flex-col lg:flex-row lg:items-center justify-between gap-6 mb-8 md:mb-12">
          <div className="flex items-center gap-4">
            <div className="w-12 h-12 sm:w-16 sm:h-16 rounded-2xl bg-zinc-900 flex items-center justify-center text-white shadow-xl">
              <UserIcon size={28} className="sm:hidden" />
              <UserIcon size={32} className="hidden sm:block" />
            </div>
            <div>
              <h1 className="text-xl sm:text-2xl md:text-3xl font-bold text-zinc-900">Welcome, {user.name}</h1>
              <p className="text-sm sm:text-base text-zinc-500 font-medium">Employee Dashboard</p>
            </div>
          </div>
          <div className="flex flex-wrap gap-3">
            <button 
              onClick={() => { setCameraMode('in'); setShowCamera(true); }}
              disabled={!!todayRecord?.timeIn}
              className="flex-1 sm:flex-none flex items-center justify-center gap-2 bg-emerald-600 text-white px-4 sm:px-8 py-3 sm:py-4 rounded-2xl hover:bg-emerald-700 transition-all shadow-lg disabled:opacity-50 font-bold text-sm sm:text-base"
            >
              <Clock size={18} />
              Time In
            </button>
            <button 
              onClick={() => { setCameraMode('out'); setShowCamera(true); }}
              disabled={!todayRecord?.timeIn || !!todayRecord?.timeOut}
              className="flex-1 sm:flex-none flex items-center justify-center gap-2 bg-blue-600 text-white px-4 sm:px-8 py-3 sm:py-4 rounded-2xl hover:bg-blue-700 transition-all shadow-lg disabled:opacity-50 font-bold text-sm sm:text-base"
            >
              <LogOut size={18} />
              Time Out
            </button>
            <button onClick={onLogout} className="p-3 sm:p-4 bg-white border border-zinc-200 text-zinc-500 rounded-2xl hover:bg-zinc-50 transition-all shadow-sm">
              <LogOut size={18} />
            </button>
          </div>
        </header>

        <div className="grid grid-cols-1 lg:grid-cols-3 gap-8">
          {/* Stats */}
          <div className="lg:col-span-1 flex flex-col gap-6">
            <div className="bg-white p-6 sm:p-8 rounded-3xl shadow-sm border border-zinc-200">
              <h3 className="text-xs font-bold text-zinc-400 uppercase tracking-wider mb-6">Today's Status</h3>
              <div className="space-y-6">
                <div className="flex items-center justify-between">
                  <div className="flex items-center gap-3">
                    <div className="w-10 h-10 bg-emerald-50 text-emerald-600 rounded-xl flex items-center justify-center">
                      <Clock size={20} />
                    </div>
                    <span className="font-medium text-zinc-600">Time In</span>
                  </div>
                  <span className="font-bold text-zinc-900">{todayRecord?.timeIn ? format(parseISO(todayRecord.timeIn), 'hh:mm a') : '--:--'}</span>
                </div>
                <div className="flex items-center justify-between">
                  <div className="flex items-center gap-3">
                    <div className="w-10 h-10 bg-blue-50 text-blue-600 rounded-xl flex items-center justify-center">
                      <LogOut size={20} />
                    </div>
                    <span className="font-medium text-zinc-600">Time Out</span>
                  </div>
                  <span className="font-bold text-zinc-900">{todayRecord?.timeOut ? format(parseISO(todayRecord.timeOut), 'hh:mm a') : '--:--'}</span>
                </div>
              </div>
            </div>

            <div className="bg-zinc-900 p-6 sm:p-8 rounded-3xl shadow-xl text-white">
              <h3 className="text-sm font-bold text-zinc-400 uppercase tracking-wider mb-4">Attendance Summary</h3>
              <div className="text-4xl font-bold mb-2">{records.length}</div>
              <p className="text-zinc-400 text-sm">Total days recorded this month</p>
            </div>
          </div>

          {/* Main Views */}
          <div className="lg:col-span-2 flex flex-col gap-6">
            <div className="flex bg-white p-1.5 rounded-2xl shadow-sm border border-zinc-200 w-full sm:w-fit">
              <button 
                onClick={() => setView('calendar')}
                className={cn("flex-1 sm:flex-none flex items-center justify-center gap-2 px-6 py-3 rounded-xl transition-all font-bold text-sm", view === 'calendar' ? "bg-zinc-900 text-white shadow-lg" : "text-zinc-500 hover:bg-zinc-50")}
              >
                <CalendarIcon size={18} />
                Calendar
              </button>
              <button 
                onClick={() => setView('table')}
                className={cn("flex-1 sm:flex-none flex items-center justify-center gap-2 px-6 py-3 rounded-xl transition-all font-bold text-sm", view === 'table' ? "bg-zinc-900 text-white shadow-lg" : "text-zinc-500 hover:bg-zinc-50")}
              >
                <TableIcon size={18} />
                Table
              </button>
            </div>

            <AnimatePresence mode="wait">
              {view === 'calendar' ? (
                <motion.div key="calendar" initial={{ opacity: 0, x: 20 }} animate={{ opacity: 1, x: 0 }} exit={{ opacity: 0, x: -20 }}>
                  <Calendar records={records} userId={user.id} />
                </motion.div>
              ) : (
                <motion.div key="table" initial={{ opacity: 0, x: 20 }} animate={{ opacity: 1, x: 0 }} exit={{ opacity: 0, x: -20 }}>
                  <div className="bg-white rounded-2xl shadow-sm border border-zinc-200 overflow-x-auto scrollbar-hide">
                    <table className="w-full text-left min-w-[800px]">
                      <thead className="bg-zinc-50 border-b border-zinc-200">
                        <tr>
                          <th className="px-6 py-4 text-[10px] font-bold text-zinc-400 uppercase tracking-wider">Date</th>
                          <th className="px-6 py-4 text-[10px] font-bold text-zinc-400 uppercase tracking-wider">Time In</th>
                          <th className="px-6 py-4 text-[10px] font-bold text-zinc-400 uppercase tracking-wider">Time Out</th>
                          <th className="px-6 py-4 text-[10px] font-bold text-zinc-400 uppercase tracking-wider">Status</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-zinc-100">
                        {records.map(r => (
                          <tr key={r.id} className="hover:bg-zinc-50 transition-colors">
                            <td className="px-6 py-4 font-medium text-zinc-900 text-sm">{r.date}</td>
                            <td className="px-6 py-4 text-zinc-500 text-sm">{r.timeIn ? format(parseISO(r.timeIn), 'hh:mm a') : '-'}</td>
                            <td className="px-6 py-4 text-zinc-500 text-sm">{r.timeOut ? format(parseISO(r.timeOut), 'hh:mm a') : '-'}</td>
                            <td className="px-6 py-4">
                              <span className={cn(
                                "px-3 py-1 rounded-full text-[10px] font-bold uppercase tracking-wider",
                                r.status === 'present' ? "bg-emerald-100 text-emerald-700" : "bg-amber-100 text-amber-700"
                              )}>
                                {r.status}
                              </span>
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </motion.div>
              )}
            </AnimatePresence>
          </div>
        </div>
      </div>

      {/* Missed Record Notification */}
      {missedRecord && (
        <div className="fixed inset-0 bg-black/50 backdrop-blur-sm flex items-center justify-center z-50 p-6">
          <motion.div 
            initial={{ scale: 0.9, opacity: 0 }}
            animate={{ scale: 1, opacity: 1 }}
            className="bg-white rounded-3xl shadow-2xl max-w-md w-full p-8"
          >
            <div className="flex flex-col items-center text-center gap-6">
              <div className="w-20 h-20 bg-amber-50 text-amber-500 rounded-full flex items-center justify-center">
                <AlertCircle size={48} />
              </div>
              <div>
                <h3 className="text-2xl font-bold text-zinc-900">Missed Time Out</h3>
                <p className="text-zinc-500 mt-2">You missed timing out on <strong>{missedRecord.date}</strong>. Please provide a reason for this gap.</p>
              </div>
              <textarea 
                value={gapReason}
                onChange={e => setGapReason(e.target.value)}
                className="w-full px-4 py-3 rounded-2xl border border-zinc-200 focus:ring-2 focus:ring-zinc-900 focus:border-transparent outline-none transition-all min-h-[120px]"
                placeholder="Explain the reason for the missing record..."
              />
              <button 
                onClick={handleSubmitGap}
                disabled={!gapReason}
                className="w-full bg-zinc-900 text-white py-4 rounded-2xl font-bold hover:bg-zinc-800 transition-all shadow-lg disabled:opacity-50"
              >
                Submit Reason
              </button>
            </div>
          </motion.div>
        </div>
      )}

      {/* Camera Modal */}
      {showCamera && (
        <div className="fixed inset-0 bg-black/50 backdrop-blur-sm flex items-center justify-center z-50 p-6">
          <motion.div 
            initial={{ scale: 0.9, opacity: 0 }}
            animate={{ scale: 1, opacity: 1 }}
            className="bg-white rounded-3xl shadow-2xl max-w-2xl w-full p-8"
          >
            <div className="flex items-center justify-between mb-8">
              <h3 className="text-2xl font-bold text-zinc-900">Face Verification - Time {cameraMode === 'in' ? 'In' : 'Out'}</h3>
              <button onClick={() => setShowCamera(false)} className="p-2 hover:bg-zinc-100 rounded-full">
                <XCircle size={24} className="text-zinc-400" />
              </button>
            </div>
            <Camera onCapture={handleTimeAction} registeredFace={user.faceImage} />
          </motion.div>
        </div>
      )}
    </div>
  );
};

const Login = ({ onLogin }: { onLogin: (user: User) => void }) => {
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState('');
  const [loginMode, setLoginMode] = useState<'google' | 'email' | 'username'>('google');
  const [email, setEmail] = useState('');
  const [username, setUsername] = useState('');
  const [password, setPassword] = useState('');

  const handleUsernameLogin = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!username || !password) return;
    setLoading(true);
    setError('');
    try {
      // 1. Find user by username using the public usernames collection
      const usernameLower = username.toLowerCase();
      let usernameDoc = await getDoc(doc(db, 'usernames', usernameLower));
      
      let userEmail = '';
      let userId = '';

      if (!usernameDoc.exists()) {
        // Fallback: Search the users collection directly if mapping is missing
        const q = query(collection(db, 'users'), where('username', '==', usernameLower));
        const snap = await getDocs(q);
        if (!snap.empty) {
          const userData = snap.docs[0].data();
          userEmail = userData.email;
          userId = snap.docs[0].id;
          // Repair mapping for future logins
          await setDoc(doc(db, 'usernames', usernameLower), { email: userEmail, userId });
        } else {
          throw new Error('Username not found.');
        }
      } else {
        const data = usernameDoc.data();
        userEmail = data?.email;
        userId = data?.userId;
      }

      if (!userEmail) {
        throw new Error('User email not found for this username.');
      }

      // 2. Sign in with email and password
      const result = await signInWithEmailAndPassword(auth, userEmail, password);
      
      // 3. Get full user data
      const userDoc = await getDoc(doc(db, 'users', result.user.uid));
      if (userDoc.exists()) {
        onLogin({ id: userDoc.id, ...userDoc.data() } as User);
      } else {
        throw new Error('User profile not found.');
      }
    } catch (err: any) {
      console.error(err);
      if (err.message === 'Username not found.') {
        setError('Username not found.');
      } else if (err.code === 'auth/user-not-found' || err.code === 'auth/wrong-password' || err.code === 'auth/invalid-credential') {
        setError('Invalid password.');
      } else {
        setError(err.message || 'Failed to sign in. Please try again.');
      }
    } finally {
      setLoading(false);
    }
  };

  const handleGoogleLogin = async () => {
    setLoading(true);
    setError('');
    try {
      const result = await signInWithPopup(auth, googleProvider);
      const userEmail = result.user.email;
      
      // 1. Try to find by UID
      let userDoc = await getDoc(doc(db, 'users', result.user.uid));
      
      // 2. If not found, try to find by email (for pre-created users)
      if (!userDoc.exists() && userEmail) {
        // Try direct lookup with pending prefix
        const pendingId = `pending_${userEmail}`;
        const pendingDoc = await getDoc(doc(db, 'users', pendingId));
        
        if (pendingDoc.exists()) {
          const userData = pendingDoc.data();
          // Transfer data to UID-based document
          await setDoc(doc(db, 'users', result.user.uid), {
            ...userData,
            authMethod: 'google'
          });
          
          // Update username mapping to point to real UID
          if (userData.username) {
            await setDoc(doc(db, 'usernames', userData.username), { 
              email: userEmail, 
              userId: result.user.uid 
            });
          }

          // Delete the pending doc
          await deleteDoc(pendingDoc.ref);
          userDoc = await getDoc(doc(db, 'users', result.user.uid));
        } else {
          // Fallback to query if ID format changed
          const q = query(collection(db, 'users'), where('email', '==', userEmail));
          const querySnapshot = await getDocs(q);
          if (!querySnapshot.empty) {
            const pendingDoc = querySnapshot.docs[0];
            const userData = pendingDoc.data();
            await setDoc(doc(db, 'users', result.user.uid), {
              ...userData,
              authMethod: 'google'
            });

            // Update username mapping to point to real UID
            if (userData.username) {
              await setDoc(doc(db, 'usernames', userData.username), { 
                email: userEmail, 
                userId: result.user.uid 
              });
            }

            await deleteDoc(pendingDoc.ref);
            userDoc = await getDoc(doc(db, 'users', result.user.uid));
          }
        }
      }
      
      if (userDoc.exists()) {
        const userData = { id: userDoc.id, ...userDoc.data() } as User;
        onLogin(userData);
      } else if (userEmail === "isaiahnoelsalazar474@gmail.com") {
        // Bootstrap admin: Auto-create document
        const adminData: User = {
          id: result.user.uid,
          email: userEmail,
          username: 'admin',
          role: 'admin',
          name: 'System Administrator',
          authMethod: 'google',
          createdAt: new Date().toISOString()
        };
        await setDoc(doc(db, 'users', result.user.uid), adminData);
        onLogin({ id: result.user.uid, ...adminData });
      } else {
        setError('Access denied. Please contact your administrator to register your account.');
        await signOut(auth);
      }
    } catch (err) {
      console.error(err);
      setError('Failed to sign in with Google.');
    } finally {
      setLoading(false);
    }
  };

  const handleEmailLogin = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!email || !password) return;
    setLoading(true);
    setError('');
    try {
      const result = await signInWithEmailAndPassword(auth, email, password);
      const userEmail = result.user.email;
      
      // Try to find by UID
      let userDoc = await getDoc(doc(db, 'users', result.user.uid));
      
      // If not found, try to find by email (for pre-created users)
      if (!userDoc.exists() && userEmail) {
        const q = query(collection(db, 'users'), where('email', '==', userEmail));
        const querySnapshot = await getDocs(q);
        if (!querySnapshot.empty) {
          const pendingDoc = querySnapshot.docs[0];
          const userData = pendingDoc.data();
          await setDoc(doc(db, 'users', result.user.uid), userData);
          userDoc = await getDoc(doc(db, 'users', result.user.uid));
        }
      }
      
      if (userDoc.exists()) {
        const userData = { id: userDoc.id, ...userDoc.data() } as User;
        onLogin(userData);
      } else if (userEmail === "isaiahnoelsalazar474@gmail.com") {
        // Bootstrap admin: Auto-create document
        const adminData: User = {
          id: result.user.uid,
          email: userEmail,
          username: 'admin',
          role: 'admin',
          name: 'System Administrator',
          authMethod: 'email',
          createdAt: new Date().toISOString()
        };
        await setDoc(doc(db, 'users', result.user.uid), adminData);
        onLogin({ id: result.user.uid, ...adminData });
      } else {
        setError('Access denied. Please contact your administrator.');
        await signOut(auth);
      }
    } catch (err: any) {
      console.error(err);
      if (err.code === 'auth/user-not-found' || err.code === 'auth/wrong-password' || err.code === 'auth/invalid-credential') {
        setError('Invalid email or password.');
      } else {
        setError('Failed to sign in. Please try again.');
      }
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className="min-h-screen bg-zinc-50 flex items-center justify-center p-4 sm:p-6">
      <motion.div 
        initial={{ opacity: 0, y: 20 }}
        animate={{ opacity: 1, y: 0 }}
        className="bg-white p-6 sm:p-10 rounded-3xl shadow-2xl max-w-md w-full border border-zinc-200"
      >
        <div className="flex flex-col items-center text-center gap-4 sm:gap-6 mb-8 sm:mb-10">
          <div className="w-16 h-16 sm:w-20 sm:h-20 bg-zinc-900 rounded-3xl flex items-center justify-center text-white shadow-2xl">
            <Clock size={40} className="sm:hidden" />
            <Clock size={48} className="hidden sm:block" />
          </div>
          <div>
            <h1 className="text-2xl sm:text-3xl font-bold text-zinc-900">Fastock Attendance</h1>
            <p className="text-sm sm:text-base text-zinc-500 mt-2">Sign in to your account</p>
          </div>
        </div>

        <div className="space-y-6">
          {error && (
            <div className="p-4 bg-red-50 text-red-600 rounded-2xl text-sm font-medium flex items-center gap-3">
              <AlertCircle size={18} />
              {error}
            </div>
          )}

          <div className="flex bg-zinc-100 p-1 rounded-2xl mb-6 sm:mb-8">
            <button 
              onClick={() => setLoginMode('google')}
              className={cn("flex-1 py-2 rounded-xl text-sm font-bold transition-all", loginMode === 'google' ? "bg-white text-zinc-900 shadow-sm" : "text-zinc-500")}
            >
              Google
            </button>
            <button 
              onClick={() => setLoginMode('email')}
              className={cn("flex-1 py-2 rounded-xl text-sm font-bold transition-all", loginMode === 'email' ? "bg-white text-zinc-900 shadow-sm" : "text-zinc-500")}
            >
              Email
            </button>
            <button 
              onClick={() => setLoginMode('username')}
              className={cn("flex-1 py-2 rounded-xl text-sm font-bold transition-all", loginMode === 'username' ? "bg-white text-zinc-900 shadow-sm" : "text-zinc-500")}
            >
              Username
            </button>
          </div>

          {loginMode === 'google' ? (
            <button 
              onClick={handleGoogleLogin}
              disabled={loading}
              className="w-full bg-zinc-900 text-white py-4 rounded-2xl font-bold hover:bg-zinc-800 transition-all shadow-lg flex items-center justify-center gap-3 disabled:opacity-50"
            >
              {loading ? (
                <div className="w-5 h-5 border-2 border-white border-t-transparent rounded-full animate-spin" />
              ) : (
                <LogIn size={20} />
              )}
              Sign in with Google
            </button>
          ) : loginMode === 'email' ? (
            <form onSubmit={handleEmailLogin} className="space-y-4">
              <div>
                <label className="block text-xs font-bold text-zinc-400 uppercase tracking-wider mb-2">Email Address</label>
                <input 
                  type="email" 
                  value={email}
                  onChange={e => setEmail(e.target.value)}
                  className="w-full px-4 py-3 rounded-xl border border-zinc-200 focus:ring-2 focus:ring-zinc-900 focus:border-transparent outline-none transition-all"
                  placeholder="name@company.com"
                  required
                />
              </div>
              <div>
                <label className="block text-xs font-bold text-zinc-400 uppercase tracking-wider mb-2">Password</label>
                <input 
                  type="password" 
                  value={password}
                  onChange={e => setPassword(e.target.value)}
                  className="w-full px-4 py-3 rounded-xl border border-zinc-200 focus:ring-2 focus:ring-zinc-900 focus:border-transparent outline-none transition-all"
                  placeholder="••••••••"
                  required
                />
              </div>
              <button 
                type="submit"
                disabled={loading}
                className="w-full bg-zinc-900 text-white py-4 rounded-2xl font-bold hover:bg-zinc-800 transition-all shadow-lg flex items-center justify-center gap-3 disabled:opacity-50"
              >
                {loading ? (
                  <div className="w-5 h-5 border-2 border-white border-t-transparent rounded-full animate-spin" />
                ) : (
                  <LogIn size={20} />
                )}
                Sign in with Email
              </button>
            </form>
          ) : (
            <form onSubmit={handleUsernameLogin} className="space-y-4">
              <div>
                <label className="block text-xs font-bold text-zinc-400 uppercase tracking-wider mb-2">Username / User ID</label>
                <input 
                  type="text" 
                  value={username}
                  onChange={e => setUsername(e.target.value)}
                  className="w-full px-4 py-3 rounded-xl border border-zinc-200 focus:ring-2 focus:ring-zinc-900 focus:border-transparent outline-none transition-all"
                  placeholder="johndoe123"
                  required
                />
              </div>
              <div>
                <label className="block text-xs font-bold text-zinc-400 uppercase tracking-wider mb-2">Password</label>
                <input 
                  type="password" 
                  value={password}
                  onChange={e => setPassword(e.target.value)}
                  className="w-full px-4 py-3 rounded-xl border border-zinc-200 focus:ring-2 focus:ring-zinc-900 focus:border-transparent outline-none transition-all"
                  placeholder="••••••••"
                  required
                />
              </div>
              <button 
                type="submit"
                disabled={loading}
                className="w-full bg-zinc-900 text-white py-4 rounded-2xl font-bold hover:bg-zinc-800 transition-all shadow-lg flex items-center justify-center gap-3 disabled:opacity-50"
              >
                {loading ? (
                  <div className="w-5 h-5 border-2 border-white border-t-transparent rounded-full animate-spin" />
                ) : (
                  <LogIn size={20} />
                )}
                Sign in with Username
              </button>
            </form>
          )}
        </div>

        <div className="mt-10 pt-10 border-t border-zinc-100 text-center">
          <p className="text-xs text-zinc-400">Secure access powered by Firebase</p>
        </div>
      </motion.div>
    </div>
  );
};

interface ErrorBoundaryProps {
  children: React.ReactNode;
}

interface ErrorBoundaryState {
  hasError: boolean;
  errorInfo: string | null;
}

class ErrorBoundary extends Component<ErrorBoundaryProps, ErrorBoundaryState> {
  constructor(props: ErrorBoundaryProps) {
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
      let displayMessage = "Something went wrong.";
      try {
        const parsed = JSON.parse(this.state.errorInfo || "");
        if (parsed.error && parsed.operationType) {
          displayMessage = `Firestore Error: ${parsed.operationType} on ${parsed.path || 'unknown path'} failed. ${parsed.error}`;
        }
      } catch (e) {
        displayMessage = this.state.errorInfo || "An unexpected error occurred.";
      }

      return (
        <div className="min-h-screen bg-zinc-50 flex items-center justify-center p-6">
          <div className="bg-white p-10 rounded-3xl shadow-2xl max-w-md w-full text-center border border-zinc-200">
            <div className="w-20 h-20 bg-red-50 text-red-500 rounded-full flex items-center justify-center mx-auto mb-6">
              <AlertCircle size={48} />
            </div>
            <h2 className="text-2xl font-bold text-zinc-900 mb-4">Application Error</h2>
            <p className="text-zinc-500 mb-8">{displayMessage}</p>
            <button 
              onClick={() => window.location.reload()}
              className="w-full bg-zinc-900 text-white py-4 rounded-2xl font-bold hover:bg-zinc-800 transition-all shadow-lg"
            >
              Reload Application
            </button>
          </div>
        </div>
      );
    }

    return this.props.children;
  }
}

export default function App() {
  const [user, setUser] = useState<User | null>(null);
  const [loading, setLoading] = useState(true);

  useEffect(() => {
    const unsub = onAuthStateChanged(auth, async (firebaseUser) => {
      try {
        if (firebaseUser) {
          const userDoc = await getDoc(doc(db, 'users', firebaseUser.uid));
          if (userDoc.exists()) {
            const userData = userDoc.data() as User;
            setUser({ id: userDoc.id, ...userData } as User);
          } else if (firebaseUser.email === "isaiahnoelsalazar474@gmail.com") {
            // Bootstrap admin: Auto-create document if it doesn't exist
            const adminData: User = {
              id: firebaseUser.uid,
              email: firebaseUser.email,
              username: 'admin',
              role: 'admin',
              name: 'System Administrator',
              authMethod: 'google',
              createdAt: new Date().toISOString()
            };
            await setDoc(doc(db, 'users', firebaseUser.uid), adminData);
            setUser({ id: firebaseUser.uid, ...adminData });
          } else {
            setUser(null);
          }
        } else {
          setUser(null);
        }
      } catch (err) {
        console.error("Auth state change error:", err);
        setUser(null);
      } finally {
        setLoading(false);
      }
    });
    return () => unsub();
  }, []);

  const handleLogout = async () => {
    await signOut(auth);
    setUser(null);
  };

  if (loading) {
    return (
      <div className="min-h-screen bg-zinc-50 flex items-center justify-center">
        <div className="w-12 h-12 border-4 border-zinc-900 border-t-transparent rounded-full animate-spin" />
      </div>
    );
  }

  if (!user) return <Login onLogin={setUser} />;

  return (
    <ErrorBoundary>
      {user.role === 'admin' ? (
        <AdminDashboard user={user} onLogout={handleLogout} />
      ) : (
        <EmployeeDashboard user={user} onLogout={handleLogout} />
      )}
    </ErrorBoundary>
  );
}
