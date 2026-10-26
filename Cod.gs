<!doctype html>
<html lang="th">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>ระบบเช็คชื่อนักเรียน</title>
  <script src="https://cdn.jsdelivr.net/npm/sweetalert2@11"></script>
  <link href="https://fonts.googleapis.com/css2?family=Sarabun:wght@300;400;600;700&display=swap" rel="stylesheet">
  <script src="https://cdn.tailwindcss.com"></script>
  <style>
    body { font-family: 'Sarabun', sans-serif; }
    /* ... */
  </style>
</head>
<body>
  <!-- ...same UI... -->
  
  <script>
    // Local storage data layer
    const STORAGE_KEY = 'attendance_system_v2';
    
    let allData = {
      students: [],
      subjects: [],
      attendance: []
    };
    
    // Load from localStorage
    function loadFromStorage() {
      try {
        const saved = localStorage.getItem(STORAGE_KEY);
        if (saved) {
          const parsed = JSON.parse(saved);
          allData = {
            students: parsed.students || [],
            subjects: parsed.subjects || [],
            attendance: parsed.attendance || []
          };
        }
      } catch (e) {
        console.error('Load error:', e);
      }
    }
    
    // Save to localStorage
    function saveToStorage() {
      try {
        localStorage.setItem(STORAGE_KEY, JSON.stringify(allData));
        return true;
      } catch (e) {
        console.error('Save error:', e);
        showErrorMessage('ไม่สามารถบันทึกข้อมูลได้: ' + e.message);
        return false;
      }
    }
    
    // Export JSON
    function exportData() {
      const data = JSON.stringify(allData, null, 2);
      const blob = new Blob([data], { type: 'application/json' });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `attendance-backup-${todayThaiStr()}.json`;
      a.click();
      URL.revokeObjectURL(url);
    }
    
    // Import JSON
    async function importData() {
      const input = document.createElement('input');
      input.type = 'file';
      input.accept = '.json';
      input.onchange = async (e) => {
        const file = e.target.files[0];
        if (!file) return;
        const text = await file.text();
        try {
          const imported = JSON.parse(text);
          const result = await Swal.fire({
            title: 'ยืนยันการนำเข้า',
            html: `จะแทนที่ข้อมูลปัจจุบันด้วย:<br>
              👨‍🎓 นักเรียน: ${imported.students?.length || 0} คน<br>
              📚 วิชา: ${imported.subjects?.length || 0} วิชา<br>
              ✅ การเช็คชื่อ: ${imported.attendance?.length || 0} รายการ<br><br>
              <strong class="text-red-600">ข้อมูลเดิมจะหายไป!</strong>`,
            icon: 'warning',
            showCancelButton: true,
            confirmButtonText: 'นำเข้า',
            cancelButtonText: 'ยกเลิก',
            confirmButtonColor: '#dc2626'
          });
          if (!result.isConfirmed) return;
          
          allData = {
            students: imported.students || [],
            subjects: imported.subjects || [],
            attendance: imported.attendance || []
          };
          saveToStorage();
          refreshUI();
          Swal.fire('สำเร็จ', 'นำเข้าข้อมูลเรียบร้อย', 'success');
        } catch (err) {
          Swal.fire('ผิดพลาด', 'ไฟล์ไม่ถูกต้อง: ' + err.message, 'error');
        }
      };
      input.click();
    }
    
    // Clear all data
    async function clearAllData() {
      const result = await Swal.fire({
        title: 'ลบข้อมูลทั้งหมด?',
        text: 'ข้อมูลทั้งหมดจะหายไป ไม่สามารถกู้คืนได้',
        icon: 'warning',
        showCancelButton: true,
        confirmButtonText: 'ลบทั้งหมด',
        cancelButtonText: 'ยกเลิก',
        confirmButtonColor: '#dc2626'
      });
      if (!result.isConfirmed) return;
      allData = { students: [], subjects: [], attendance: [] };
      saveToStorage();
      refreshUI();
      Swal.fire('สำเร็จ', 'ลบข้อมูลทั้งหมดแล้ว', 'success');
    }
    
    // helper functions
    function todayThaiStr() {
      const now = new Date();
      const b = new Date(now.getTime() + 7 * 3600 * 1000);
      return b.toISOString().split('T')[0];
    }
    
    function esc(s) {
      return String(s ?? '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#39;');
    }
    
    // ... rest of the functions
  </script>
</body>
</html>