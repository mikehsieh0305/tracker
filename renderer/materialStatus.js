// renderer/materialStatus.js
export class MaterialStatus {
  constructor(row) {
    this.area           = row['區域'] || '';
    this.capacityKw     = Number(row['容量(Kw)'] || 0);

    this.pileRate       = Number(row['基樁發料完成率'] || row['基樁完成率'] || 0);
    this.steelMainRate  = Number(row['鋼構-大料發料完成率'] || 0);
    this.steelSubRate   = Number(row['鋼構-小料發料完成率'] || 0);
    this.moduleRackRate = Number(row['模組架完成率'] || row['模組支架發料完成率'] || 0);
    this.moduleRate     = Number(row['模組發料完成率'] || 0);

    this.materialStatus = '';
  }

  updateStatus() {
    const steelMainZero  = this.steelMainRate === 0;
    const steelSubZero   = this.steelSubRate === 0;
    const moduleRackZero = this.moduleRackRate === 0;
    const moduleZero     = this.moduleRate === 0;

    const zeroCount = [steelMainZero, steelSubZero, moduleRackZero, moduleZero].filter(Boolean).length;

    if (zeroCount === 4) {
      this.materialStatus = '嚴重缺料（鋼構與模組皆未發料）';
    } else if (zeroCount >= 2) {
      this.materialStatus = '多項缺料（大料/小料/模組架/模組部分未到）';
    } else if (zeroCount === 1) {
      if (steelMainZero)  this.materialStatus = '缺鋼構大料';
      else if (steelSubZero)   this.materialStatus = '缺鋼構小料';
      else if (moduleRackZero) this.materialStatus = '缺模組架';
      else if (moduleZero)     this.materialStatus = '缺模組';
    } else {
      this.materialStatus = '材料正常';
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
      '材料狀態': this.materialStatus
    };
  }
}
