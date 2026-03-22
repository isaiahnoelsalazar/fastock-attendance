import { GoogleGenAI } from "@google/genai";

export interface User {
  id: string;
  name: string;
  email: string;
  username: string;
  role: 'admin' | 'employee';
  authMethod?: 'google' | 'email';
  employeeType?: 'regular' | 'intern';
  requiredHours?: number;
  faceImage?: string; // Base64
  createdAt?: string;
}

export interface AttendanceRecord {
  id: string;
  userId: string;
  date: string; // YYYY-MM-DD
  timeIn?: string; // ISO string
  timeOut?: string; // ISO string
  status: 'present' | 'absent' | 'missed';
}

export interface GapReason {
  id: string;
  userId: string;
  recordId: string;
  reason: string;
  status: 'pending' | 'approved' | 'denied';
  adminComment?: string;
}

const ai = new GoogleGenAI({ apiKey: process.env.GEMINI_API_KEY || "" });

export async function verifyFace(registeredFace: string, currentFace: string): Promise<boolean> {
  try {
    const model = ai.models.generateContent({
      model: "gemini-3-flash-preview",
      contents: [
        {
          parts: [
            { text: "Do these two images show the same person? Reply with only 'YES' or 'NO'." },
            { inlineData: { mimeType: "image/jpeg", data: registeredFace.split(',')[1] } },
            { inlineData: { mimeType: "image/jpeg", data: currentFace.split(',')[1] } }
          ]
        }
      ]
    });
    const response = await model;
    return response.text.trim().toUpperCase() === 'YES';
  } catch (error) {
    console.error("Face verification error:", error);
    return false;
  }
}
