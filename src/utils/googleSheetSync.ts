import * as XLSX from 'xlsx';
import { Student, GoogleSheetConfig } from '../types';
import { autoRepairVietnameseText } from './vietnameseEncoding';
import { extractStructuredSheet, getRowValue } from './excelParser';

/**
 * Converts standard Google Sheet share URLs into public CSV export endpoints
 */
export function getGoogleSheetCsvUrl(url: string): string {
  const trimmed = url.trim();
  if (!trimmed) return '';

  if (
    trimmed.includes('export?format=csv') ||
    trimmed.includes('/gviz/tq?tqx=out:csv') ||
    trimmed.includes('/pub?output=csv')
  ) {
    return trimmed;
  }

  let gid = '';
  const gidMatch = trimmed.match(/[#&?]gid=([0-9]+)/);
  if (gidMatch) {
    gid = gidMatch[1];
  }

  // Published sheet: /d/e/PUB_ID/
  const pubMatch = trimmed.match(/\/d\/e\/([a-zA-Z0-9-_]+)/);
  if (pubMatch) {
    const pubId = pubMatch[1];
    return `https://docs.google.com/spreadsheets/d/e/${pubId}/pub?output=csv${gid ? `&gid=${gid}` : ''}`;
  }

  // Standard sheet: /d/SHEET_ID/
  const idMatch = trimmed.match(/\/d\/([a-zA-Z0-9-_]+)/);
  if (idMatch) {
    const sheetId = idMatch[1];
    return `https://docs.google.com/spreadsheets/d/${sheetId}/export?format=csv${gid ? `&gid=${gid}` : ''}`;
  }

  return trimmed;
}

/**
 * Sample Google Sheet URL for instant testing
 */
export const SAMPLE_GOOGLE_SHEET_URL =
  'https://docs.google.com/spreadsheets/d/1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms/edit?usp=sharing';

/**
 * Fetches online CSV from Google Sheets and parses it into Student objects
 */
export async function fetchStudentsFromGoogleSheet(sheetUrl: string): Promise<Student[]> {
  const csvUrl = getGoogleSheetCsvUrl(sheetUrl);
  if (!csvUrl) {
    throw new Error('Đường dẫn Google Sheet không hợp lệ. Vui lòng dán đúng liên kết Google Sheets.');
  }

  let csvText = '';

  // 1. Direct browser fetch
  try {
    const res = await fetch(csvUrl, { mode: 'cors' });
    if (res.ok) {
      csvText = await res.text();
    }
  } catch {
    // CORS or network fallback
  }

  // 2. Server proxy fallback if browser fetch failed
  if (!csvText) {
    try {
      const proxyRes = await fetch('/api/sync-google-sheet', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ sheetUrl: csvUrl }),
      });
      if (proxyRes.ok) {
        const proxyData = await proxyRes.json();
        if (proxyData.csvText) {
          csvText = proxyData.csvText;
        }
      }
    } catch {
      // proxy fallback error
    }
  }

  if (!csvText) {
    throw new Error(
      'Không thể tải dữ liệu từ Google Sheet. Vui lòng mở Google Sheet -> chọn "Chia sẻ" (Share) -> bật "Bất kỳ ai có liên kết đều có thể xem" (Anyone with the link can view).'
    );
  }

  const workbook = XLSX.read(csvText, { type: 'string', raw: false });
  if (!workbook.SheetNames || workbook.SheetNames.length === 0) {
    throw new Error('Không đọc được bảng dữ liệu từ Google Sheet.');
  }

  const { rows: rawRows } = extractStructuredSheet(workbook, 'auto');
  if (!rawRows || rawRows.length === 0) {
    throw new Error('Bảng tính Google Sheet trống hoặc không tìm thấy dòng dữ liệu học sinh.');
  }

  const validStudents: Student[] = rawRows
    .map((row, idx) => transformRowToStudent(row, idx))
    .filter((s): s is Student => s !== null);

  if (validStudents.length === 0) {
    throw new Error(
      'Không thể nhận diện danh sách học sinh. Kiểm tra lại hàng tiêu đề Google Sheet có các cột như "Họ và Tên", "Mã HS", "Giới tính", "Ngày sinh", "Tổ"...'
    );
  }

  return validStudents;
}

/**
 * Transform raw Google Sheet row into a clean Student object
 */
function transformRowToStudent(row: any, index: number): Student | null {
  const getVal = (keys: string[]) => getRowValue(row, keys);
  const cleanStr = (val: any) => (val ? autoRepairVietnameseText(String(val).trim()) : '');

  const rawCode = getVal(['mã hs', 'mã học sinh', 'ma hs', 'code', 'id', 'mshs', 'stt']);
  const rawName = getVal(['họ và tên', 'họ tên', 'tên', 'tên học sinh', 'name', 'full name', 'fullname']);

  if (!rawName && !rawCode) return null;

  const name = rawName ? cleanStr(rawName) : `Học sinh ${index + 1}`;
  const code = rawCode ? String(rawCode).trim() : `GVCN-${(index + 1).toString().padStart(2, '0')}`;

  const rawGender = getVal(['giới tính', 'gioi tinh', 'phái', 'gender', 'sex']);
  const isFemale =
    rawGender &&
    (String(rawGender).toLowerCase().includes('nữ') ||
      String(rawGender).toLowerCase().includes('nu') ||
      String(rawGender).toLowerCase() === 'f');
  const gender = isFemale ? 'Nữ' : 'Nam';

  const rawDob = getVal(['ngày sinh', 'ngay sinh', 'dob', 'birthday', 'birth date']);
  const dob = rawDob ? String(rawDob).trim() : '15/03/2008';

  const rawGroup = getVal(['tổ', 'to', 'group', 'tổ thi đua']);
  let group: 1 | 2 | 3 | 4 = 1;
  if (rawGroup) {
    const parsedG = parseInt(String(rawGroup).replace(/[^0-9]/g, ''), 10);
    if (parsedG >= 1 && parsedG <= 4) group = parsedG as any;
  }

  const rawPhone = getVal(['sđt', 'sdt', 'số điện thoại', 'điện thoại', 'phone', 'tel']);
  const phone = rawPhone ? String(rawPhone).trim() : '0912000000';

  const rawEmail = getVal(['email', 'thư điện tử', 'mail']);
  const email = rawEmail ? String(rawEmail).trim() : `${code.toLowerCase()}@gvcn2027.edu.vn`;

  const rawAddress = getVal(['địa chỉ', 'dia chi', 'address', 'nơi ở']);
  const address = rawAddress ? cleanStr(rawAddress) : 'Hải Phòng';

  const rawStrengths = getVal(['sở trường năng khiếu', 'sở trường', 'năng khiếu', 'strengths']);
  const strengths = rawStrengths ? cleanStr(rawStrengths) : 'Toán học & Khoa học Tự nhiên';

  const rawCareer = getVal(['định hướng nghề nghiệp', 'định hướng', 'nguyện vọng', 'career']);
  const careerAspiration = rawCareer ? cleanStr(rawCareer) : 'Đại học Bách Khoa / Kinh Tế';

  const rawHealth = getVal(['ghi chú sức khỏe', 'sức khỏe', 'suc khoe', 'health note']);
  const healthNote = rawHealth ? cleanStr(rawHealth) : 'Sức khỏe tốt';

  const rawParentName = getVal(['họ tên phụ huynh', 'phụ huynh', 'tên phụ huynh', 'parent']);
  const parentName = rawParentName ? cleanStr(rawParentName) : `Phụ huynh của ${name}`;

  const rawRel = getVal(['quan hệ', 'mối quan hệ', 'relationship']);
  let relationship: 'Bố' | 'Mẹ' | 'Người giám hộ' = 'Bố';
  if (rawRel && (String(rawRel).toLowerCase().includes('mẹ') || String(rawRel).toLowerCase().includes('me'))) {
    relationship = 'Mẹ';
  }

  const rawParentPhone = getVal(['sđt phụ huynh', 'sđt ph', 'số điện thoại phụ huynh', 'parent phone']);
  const parentPhone = rawParentPhone ? String(rawParentPhone).trim() : '0912888999';

  const rawWorkplace = getVal(['nơi công tác phụ huynh', 'nơi công tác', 'workplace']);
  const workplace = rawWorkplace ? cleanStr(rawWorkplace) : 'Hải Phòng';

  const rawGpa = getVal(['đtb khối a', 'đtb', 'điểm tb', 'gpa', 'dtb']);
  const gpa = rawGpa ? Number(parseFloat(String(rawGpa)).toFixed(2)) : 8.5;

  const rawConduct = getVal(['hạnh kiểm', 'rèn luyện', 'conduct']);
  let conductRating: 'Tốt' | 'Khá' | 'Trung bình' | 'Yếu' = 'Tốt';
  if (rawConduct && (String(rawConduct).toLowerCase().includes('khá') || String(rawConduct).toLowerCase().includes('kha'))) {
    conductRating = 'Khá';
  }

  const defaultAvatar =
    gender === 'Nữ'
      ? `https://images.unsplash.com/photo-1544005313-94ddf0286df2?w=150&auto=format&fit=crop&q=80`
      : `https://images.unsplash.com/photo-1539571696357-5a69c17a67c6?w=150&auto=format&fit=crop&q=80`;

  const id = code.toLowerCase().replace(/[^a-z0-9]/g, '-') || `std-${Date.now()}-${index}`;

  return {
    id,
    code,
    name,
    gender,
    dob,
    group,
    avatar: defaultAvatar,
    phone,
    email,
    address,
    strengths,
    careerAspiration,
    healthNote,
    emergencyContact: {
      parentName,
      relationship,
      phone: parentPhone,
      workplace,
    },
    grades: {
      math: { tx1: gpa, tx2: gpa, gk: gpa, ck: gpa, avg: gpa },
      physics: { tx1: gpa, tx2: gpa, gk: gpa, ck: gpa, avg: gpa },
      chemistry: { tx1: gpa, tx2: gpa, gk: gpa, ck: gpa, avg: gpa },
      biology: { tx1: 8.5, tx2: 8.5, gk: 8.5, ck: 8.5, avg: 8.5 },
      literature: { tx1: 8.0, tx2: 8.0, gk: 8.0, ck: 8.0, avg: 8.0 },
      english: { tx1: 8.5, tx2: 8.5, gk: 8.5, ck: 8.5, avg: 8.5 },
      gpa,
    },
    progressHistory: [
      { period: 'Tháng 9', math: Math.max(0, gpa - 0.5), physics: Math.max(0, gpa - 0.4), chemistry: Math.max(0, gpa - 0.6) },
      { period: 'Giữa HK1', math: Math.max(0, gpa - 0.2), physics: Math.max(0, gpa - 0.1), chemistry: Math.max(0, gpa - 0.3) },
      { period: 'Cuối HK1', math: gpa, physics: gpa, chemistry: gpa },
    ],
    conductScore: conductRating === 'Tốt' ? 100 : 85,
    conductRating,
    violationsCount: 0,
    commendationsCount: 1,
    absenceCount: 0,
    violations: [],
    commendations: ['Gia nhập tập thể lớp'],
  };
}
