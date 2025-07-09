#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Laser Point Activity Data Corrector
修正激光点活动数据中的问题并生成清理后的数据
"""

import re
from typing import List, Dict, Tuple
from dataclasses import dataclass
from datetime import datetime

@dataclass
class TimeSegment:
    """时间段数据类"""
    name: str
    start_time: float
    end_time: float
    
    @property
    def duration(self) -> float:
        return self.end_time - self.start_time

@dataclass
class LaserInactiveRange:
    """激光点不活跃范围数据类"""
    start_time: float
    end_time: float
    
    @property
    def duration(self) -> float:
        return self.end_time - self.start_time

class LaserDataCorrector:
    """激光数据修正器"""
    
    def __init__(self):
        self.segments: List[TimeSegment] = []
        self.original_inactive_ranges: List[LaserInactiveRange] = []
        self.corrected_inactive_ranges: List[LaserInactiveRange] = []
        self.corrections_applied = []
    
    def parse_original_data(self, data_text: str):
        """解析原始数据"""
        lines = data_text.strip().split('\n')
        seen_inactive_ranges = set()
        
        for line in lines:
            line = line.strip()
            if not line:
                continue
                
            # 解析段落时间范围
            segment_match = re.search(r'段落(\d+):\s*时间范围\s+([\d.]+)s\s*-\s*([\d.]+)s', line)
            if segment_match:
                segment_num = int(segment_match.group(1))
                start_time = float(segment_match.group(2))
                end_time = float(segment_match.group(3))
                segment = TimeSegment(f"段落{segment_num}", start_time, end_time)
                self.segments.append(segment)
                continue
            
            # 解析激光点不活跃范围
            inactive_match = re.search(r'激光点不活跃:\s+([\d.]+)s\s*-\s*([\d.]+)s', line)
            if inactive_match:
                start_time = float(inactive_match.group(1))
                end_time = float(inactive_match.group(2))
                
                # 避免重复
                range_key = (start_time, end_time)
                if range_key not in seen_inactive_ranges:
                    seen_inactive_ranges.add(range_key)
                    inactive_range = LaserInactiveRange(start_time, end_time)
                    self.original_inactive_ranges.append(inactive_range)
    
    def apply_corrections(self):
        """应用所有修正"""
        if not self.segments:
            print("❌ 没有找到段落数据")
            return
        
        # 获取实际时间边界
        min_time = min(seg.start_time for seg in self.segments)
        max_time = max(seg.end_time for seg in self.segments)
        
        print(f"📊 时间边界: {min_time}s - {max_time}s")
        print(f"📋 原始不活跃范围: {len(self.original_inactive_ranges)} 个")
        
        corrected_ranges = []
        
        for i, inactive in enumerate(self.original_inactive_ranges):
            original_start = inactive.start_time
            original_end = inactive.end_time
            
            # 修正1: 交换颠倒的时间范围
            if original_start > original_end:
                corrected_start = original_end
                corrected_end = original_start
                self.corrections_applied.append(
                    f"修正{i+1}: 时间范围颠倒 {original_start}s-{original_end}s → {corrected_start}s-{corrected_end}s"
                )
            else:
                corrected_start = original_start
                corrected_end = original_end
            
            # 修正2: 限制在实际时间边界内
            if corrected_start < min_time:
                self.corrections_applied.append(
                    f"修正{i+1}: 开始时间超出下界 {corrected_start}s → {min_time}s"
                )
                corrected_start = min_time
            
            if corrected_end > max_time:
                self.corrections_applied.append(
                    f"修正{i+1}: 结束时间超出上界 {corrected_end}s → {max_time}s"
                )
                corrected_end = max_time
            
            # 修正3: 确保最小持续时间
            min_duration = 0.1
            if corrected_end - corrected_start < min_duration:
                self.corrections_applied.append(
                    f"修正{i+1}: 持续时间过短，调整为最小持续时间 {min_duration}s"
                )
                corrected_end = corrected_start + min_duration
                # 再次检查上界
                if corrected_end > max_time:
                    corrected_end = max_time
                    corrected_start = max_time - min_duration
            
            corrected_range = LaserInactiveRange(corrected_start, corrected_end)
            corrected_ranges.append(corrected_range)
        
        # 修正4: 去除重叠范围
        merged_ranges = self._merge_overlapping_ranges(corrected_ranges)
        self.corrected_inactive_ranges = merged_ranges
        
        print(f"✅ 修正后不活跃范围: {len(self.corrected_inactive_ranges)} 个")
        
        if len(corrected_ranges) != len(merged_ranges):
            self.corrections_applied.append(
                f"修正: 合并重叠范围 {len(corrected_ranges)} → {len(merged_ranges)}"
            )
    
    def _merge_overlapping_ranges(self, ranges: List[LaserInactiveRange]) -> List[LaserInactiveRange]:
        """合并重叠的时间范围"""
        if not ranges:
            return []
        
        # 按开始时间排序
        sorted_ranges = sorted(ranges, key=lambda x: x.start_time)
        merged = [sorted_ranges[0]]
        
        for current in sorted_ranges[1:]:
            last_merged = merged[-1]
            
            # 检查是否重叠
            if current.start_time <= last_merged.end_time:
                # 合并范围
                merged[-1] = LaserInactiveRange(
                    last_merged.start_time,
                    max(last_merged.end_time, current.end_time)
                )
            else:
                # 添加新范围
                merged.append(current)
        
        return merged
    
    def generate_corrected_data_format(self) -> str:
        """生成修正后的数据格式"""
        lines = []
        
        # 段落数据保持不变
        for segment in self.segments:
            lines.append(f"{segment.name}: 时间范围 {segment.start_time}s - {segment.end_time}s")
            
            # 为每个段落添加相关的激光不活跃信息
            for i, inactive in enumerate(self.corrected_inactive_ranges):
                # 检查是否与当前段落有重叠
                overlap_start = max(inactive.start_time, segment.start_time)
                overlap_end = min(inactive.end_time, segment.end_time)
                
                if overlap_start < overlap_end:
                    # 有重叠
                    lines.append(f"激光点不活跃: {inactive.start_time}s - {inactive.end_time}s 重叠于段落 {segment.start_time}s - {segment.end_time}s")
                else:
                    # 无重叠
                    lines.append(f"激光点不活跃: {inactive.start_time}s - {inactive.end_time}s 不在段落 {segment.start_time}s - {segment.end_time}s")
        
        return "\n".join(lines)
    
    def calculate_corrected_stats(self) -> Dict:
        """计算修正后的统计信息"""
        if not self.segments or not self.corrected_inactive_ranges:
            return {}
        
        total_duration = max(seg.end_time for seg in self.segments)
        total_inactive_duration = sum(r.duration for r in self.corrected_inactive_ranges)
        
        segment_stats = []
        for segment in self.segments:
            inactive_in_segment = 0.0
            
            for inactive in self.corrected_inactive_ranges:
                overlap_start = max(inactive.start_time, segment.start_time)
                overlap_end = min(inactive.end_time, segment.end_time)
                
                if overlap_start < overlap_end:
                    inactive_in_segment += overlap_end - overlap_start
            
            active_duration = segment.duration - inactive_in_segment
            activity_rate = (active_duration / segment.duration) * 100 if segment.duration > 0 else 0
            
            segment_stats.append({
                'segment': segment,
                'total_duration': segment.duration,
                'inactive_duration': inactive_in_segment,
                'active_duration': active_duration,
                'activity_rate': activity_rate
            })
        
        return {
            'total_duration': total_duration,
            'total_inactive_duration': total_inactive_duration,
            'total_active_duration': total_duration - total_inactive_duration,
            'overall_activity_rate': ((total_duration - total_inactive_duration) / total_duration) * 100,
            'segment_stats': segment_stats
        }
    
    def generate_comparison_report(self) -> str:
        """生成对比报告"""
        report = []
        report.append("=" * 60)
        report.append("激光点活动数据修正报告")
        report.append("=" * 60)
        report.append(f"修正时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        report.append("")
        
        # 应用的修正
        report.append("🔧 应用的修正:")
        if self.corrections_applied:
            for correction in self.corrections_applied:
                report.append(f"   • {correction}")
        else:
            report.append("   • 未发现需要修正的问题")
        report.append("")
        
        # 修正前后对比
        report.append("📊 修正前后对比:")
        report.append("   原始不活跃范围:")
        for i, inactive in enumerate(self.original_inactive_ranges):
            status = "❌" if inactive.start_time > inactive.end_time else "⚠️"
            report.append(f"      {i+1}. {inactive.start_time}s - {inactive.end_time}s (时长: {abs(inactive.duration):.1f}s) {status}")
        
        report.append("")
        report.append("   修正后不活跃范围:")
        for i, inactive in enumerate(self.corrected_inactive_ranges):
            report.append(f"      {i+1}. {inactive.start_time}s - {inactive.end_time}s (时长: {inactive.duration:.1f}s) ✅")
        
        report.append("")
        
        # 修正后统计
        stats = self.calculate_corrected_stats()
        if stats:
            report.append("📈 修正后活动统计:")
            report.append(f"   • 总时长: {stats['total_duration']:.1f}秒")
            report.append(f"   • 总活跃时长: {stats['total_active_duration']:.1f}秒")
            report.append(f"   • 总不活跃时长: {stats['total_inactive_duration']:.1f}秒")
            report.append(f"   • 整体活跃率: {stats['overall_activity_rate']:.1f}%")
            report.append("")
            
            report.append("   各段落活跃率:")
            for stat in stats['segment_stats']:
                segment = stat['segment']
                rate = stat['activity_rate']
                report.append(f"      • {segment.name}: {rate:.1f}% ({stat['active_duration']:.1f}s/{stat['total_duration']:.1f}s)")
        
        report.append("")
        report.append("=" * 60)
        
        return "\n".join(report)

def main():
    """主函数"""
    print("🔧 激光点活动数据修正工具")
    print("=" * 50)
    
    # 原始有问题的数据
    original_data = """
段落0: 时间范围 0.0s - 2.3s
激光点不活跃: 11.0s - 2.3s 不在段落 0.0s - 2.3s
激光点不活跃: 7.5s - 4.5s 不在段落 0.0s - 2.3s
段落1: 时间范围 2.3s - 3.5s
激光点不活跃: 11.0s - 2.3s 不在段落 2.3s - 3.5s
激光点不活跃: 7.5s - 4.5s 不在段落 2.3s - 3.5s
段落2: 时间范围 3.5s - 4.5s
激光点不活跃: 11.0s - 2.3s 不在段落 3.5s - 4.5s
激光点不活跃: 7.5s - 4.5s 不在段落 3.5s - 4.5s
段落3: 时间范围 4.5s - 5.8s
激光点不活跃: 11.0s - 2.3s 不在段落 4.5s - 5.8s
激光点不活跃: 7.5s - 4.5s 不在段落 4.5s - 5.8s
"""
    
    # 创建修正器
    corrector = LaserDataCorrector()
    
    print("🔍 解析原始数据...")
    corrector.parse_original_data(original_data)
    
    print("⚡ 应用数据修正...")
    corrector.apply_corrections()
    
    print("📋 生成修正后的数据格式...")
    corrected_data = corrector.generate_corrected_data_format()
    
    print("📊 生成对比报告...")
    comparison_report = corrector.generate_comparison_report()
    
    # 保存修正后的数据
    with open("corrected_laser_data.txt", "w", encoding="utf-8") as f:
        f.write("# 修正后的激光点活动数据\n")
        f.write("# " + "=" * 50 + "\n\n")
        f.write(corrected_data)
    
    # 保存对比报告
    with open("laser_correction_report.txt", "w", encoding="utf-8") as f:
        f.write(comparison_report)
    
    print("\n✅ 修正完成!")
    print("📁 生成的文件:")
    print("   • corrected_laser_data.txt - 修正后的数据")
    print("   • laser_correction_report.txt - 修正对比报告")
    
    print("\n" + comparison_report)
    
    print("\n" + "=" * 60)
    print("📋 修正后的清洁数据:")
    print("=" * 60)
    print(corrected_data)
    
    return corrector

if __name__ == "__main__":
    corrector = main()