// renderer/areaProgress.js
export class AreaProgress {
  constructor(row) {
    // 這裡依你現在 Excel 欄位名稱做 mapping
    this.area           = row['區域'] || '';
    this.capacityKw     = Number(row['容量(Kw)'] || 0);

    // 這些欄位名稱請對照你現場的 excel 欄頭調整
    this.pileRate       = Number(row['基樁發料完成率'] || row['基樁完成率'] || 0) * 100;
    this.steelMainRate  = Number(row['鋼構-大料發料完成率'] || 0) * 100;
    this.steelSubRate   = Number(row['鋼構-小料發料完成率'] || 0) * 100;
    this.moduleRackRate = Number(row['模組架完成率'] || row['模組支架發料完成率'] || 0) * 100;
    this.moduleRate     = Number(row['模組發料完成率'] || 0) * 100;

    this.status         = ''; // 後面 updateStatus() 會填
  }

  updateStatus() {
    console.log(this.moduleRate);
    if (this.moduleRate >= 95) {
      this.status = '接近完工';
    } else if (this.steelMainRate === 0 && this.steelSubRate === 0 && this.moduleRackRate === 0) {
      this.status = '尚未進入鋼構/模組';
    } else if (this.pileRate < 50) {
      this.status = '基樁進度落後';
    } else {
      this.status = '施工中';
    }
  }

  toPlainObject() {
    return {
      '區域': this.area,
      '容量(Kw)': this.capacityKw,
      '基樁完成率(%)': this.pileRate,
      '鋼構大料完成率(%)': this.steelMainRate,
      '鋼構小料完成率(%)': this.steelSubRate,
      '模組架完成率(%)': this.moduleRackRate,
      '模組完成率(%)': this.moduleRate,
      '狀態': this.status
    };
  }
}
