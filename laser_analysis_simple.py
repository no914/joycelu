#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Laser Point Activity Analysis Tool (Simple Version)
分析激光点活动数据，修正时间范围问题并提供详细报告
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
    
    def __str__(self):
        return f"{self.name}: {self.start_time}s - {self.end_time}s (时长: {self.duration:.1f}s)"

@dataclass
class LaserInactiveRange:
    """激光点不活跃范围数据类"""
    start_time: float
    end_time: float
    original_text: str = ""
    
    @property
    def duration(self) -> float:
        return abs(self.end_time - self.start_time)
    
    @property
    def is_valid_range(self) -> bool:
        return self.start_time <= self.end_time
    
    def fix_range(self):
        """修正时间范围顺序"""
        if not self.is_valid_range:
            self.start_time, self.end_time = self.end_time, self.start_time
    
    def __str__(self):
        status = "✓" if self.is_valid_range else "✗ (需修正)"
        return f"激光不活跃: {self.start_time}s - {self.end_time}s (时长: {self.duration:.1f}s) {status}"

class LaserActivityAnalyzer:
    """激光活动分析器"""
    
    def __init__(self):
        self.segments: List[TimeSegment] = []
        self.inactive_ranges: List[LaserInactiveRange] = []
        self.total_duration = 0.0
    
    def parse_chinese_data(self, data_text: str) -> Dict:
        """解析中文格式的激光活动数据"""
        lines = data_text.strip().split('\n')
        
        results = {
            'segments': [],
            'inactive_ranges': [],
            'issues': []
        }
        
        # 去重处理，避免重复的激光不活跃范围
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
                results['segments'].append(segment)
                self.segments.append(segment)
                continue
            
            # 解析激光点不活跃范围
            inactive_match = re.search(r'激光点不活跃:\s+([\d.]+)s\s*-\s*([\d.]+)s', line)
            if inactive_match:
                start_time = float(inactive_match.group(1))
                end_time = float(inactive_match.group(2))
                
                # 避免重复添加相同的不活跃范围
                range_key = (start_time, end_time)
                if range_key not in seen_inactive_ranges:
                    seen_inactive_ranges.add(range_key)
                    
                    inactive_range = LaserInactiveRange(start_time, end_time, line)
                    results['inactive_ranges'].append(inactive_range)
                    self.inactive_ranges.append(inactive_range)
                    
                    # 检查时间范围问题
                    if not inactive_range.is_valid_range:
                        results['issues'].append(f"时间范围错误: {line}")
        
        # 计算总时长
        if self.segments:
            self.total_duration = max(seg.end_time for seg in self.segments)
        
        return results
    
    def fix_time_ranges(self):
        """修正所有时间范围问题"""
        fixed_count = 0
        for inactive_range in self.inactive_ranges:
            if not inactive_range.is_valid_range:
                print(f"修正时间范围: {inactive_range.start_time}s-{inactive_range.end_time}s -> {inactive_range.end_time}s-{inactive_range.start_time}s")
                inactive_range.fix_range()
                fixed_count += 1
        
        print(f"修正了 {fixed_count} 个时间范围问题")
        return fixed_count
    
    def analyze_overlaps(self) -> Dict:
        """分析激光不活跃期间与段落的重叠情况"""
        analysis = {
            'overlaps': [],
            'no_overlaps': [],
            'summary': {}
        }
        
        for inactive in self.inactive_ranges:
            overlaps_found = []
            
            for segment in self.segments:
                # 检查是否有重叠
                overlap_start = max(inactive.start_time, segment.start_time)
                overlap_end = min(inactive.end_time, segment.end_time)
                
                if overlap_start < overlap_end:
                    overlap_duration = overlap_end - overlap_start
                    overlaps_found.append({
                        'segment': segment,
                        'overlap_start': overlap_start,
                        'overlap_end': overlap_end,
                        'overlap_duration': overlap_duration
                    })
            
            if overlaps_found:
                analysis['overlaps'].append({
                    'inactive_range': inactive,
                    'overlaps': overlaps_found
                })
            else:
                analysis['no_overlaps'].append(inactive)
        
        # 生成摘要
        analysis['summary'] = {
            'total_inactive_ranges': len(self.inactive_ranges),
            'ranges_with_overlaps': len(analysis['overlaps']),
            'ranges_without_overlaps': len(analysis['no_overlaps']),
            'total_segments': len(self.segments)
        }
        
        return analysis
    
    def calculate_laser_activity_stats(self) -> Dict:
        """计算激光活动统计信息"""
        if not self.segments or not self.inactive_ranges:
            return {}
        
        # 计算总的不活跃时间
        total_inactive_duration = sum(r.duration for r in self.inactive_ranges)
        
        # 计算每个段落的激光活跃时间
        segment_stats = []
        for segment in self.segments:
            inactive_in_segment = 0.0
            
            for inactive in self.inactive_ranges:
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
            'total_duration': self.total_duration,
            'total_inactive_duration': total_inactive_duration,
            'total_active_duration': self.total_duration - total_inactive_duration,
            'overall_activity_rate': ((self.total_duration - total_inactive_duration) / self.total_duration) * 100 if self.total_duration > 0 else 0,
            'segment_stats': segment_stats
        }
    
    def generate_ascii_timeline(self) -> str:
        """生成ASCII字符时间线"""
        if not self.segments:
            return "没有段落数据"
        
        timeline = []
        timeline.append("激光点活动时间线 (ASCII 版本)")
        timeline.append("=" * 50)
        
        # 时间轴长度
        timeline_length = 60
        max_time = self.total_duration
        
        # 生成时间刻度
        time_scale = []
        for i in range(0, timeline_length + 1, 10):
            time_val = (i / timeline_length) * max_time
            time_scale.append(f"{time_val:.1f}s")
        timeline.append("时间: " + " ".join(f"{t:>8}" for t in time_scale))
        timeline.append("      " + "".join("+" if i % 10 == 0 else "-" for i in range(timeline_length + 1)))
        
        # 绘制段落
        for segment in self.segments:
            start_pos = int((segment.start_time / max_time) * timeline_length)
            end_pos = int((segment.end_time / max_time) * timeline_length)
            
            line = [" "] * (timeline_length + 1)
            for i in range(start_pos, min(end_pos + 1, timeline_length + 1)):
                line[i] = "█"
            
            timeline.append(f"{segment.name:>6}: {''.join(line)} [{segment.start_time:.1f}s-{segment.end_time:.1f}s]")
        
        timeline.append("")
        
        # 绘制激光不活跃期间
        for i, inactive in enumerate(self.inactive_ranges):
            start_pos = int((inactive.start_time / max_time) * timeline_length)
            end_pos = int((inactive.end_time / max_time) * timeline_length)
            
            line = [" "] * (timeline_length + 1)
            for j in range(start_pos, min(end_pos + 1, timeline_length + 1)):
                line[j] = "X"
            
            timeline.append(f"不活跃{i+1:>2}: {''.join(line)} [{inactive.start_time:.1f}s-{inactive.end_time:.1f}s]")
        
        timeline.append("")
        timeline.append("图例: █ = 活动段落, X = 激光不活跃")
        
        return "\n".join(timeline)
    
    def generate_detailed_report(self) -> str:
        """生成详细分析报告"""
        report = []
        report.append("=" * 60)
        report.append("激光点活动数据分析报告")
        report.append("=" * 60)
        report.append(f"生成时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        report.append("")
        
        # 基本信息
        report.append("📊 基本信息:")
        report.append(f"   • 总段落数: {len(self.segments)}")
        report.append(f"   • 激光不活跃范围数: {len(self.inactive_ranges)}")
        report.append(f"   • 总时长: {self.total_duration:.1f}秒")
        report.append("")
        
        # 段落详情
        report.append("📋 段落详情:")
        for segment in self.segments:
            report.append(f"   • {segment}")
        report.append("")
        
        # 激光不活跃范围（修正前后）
        report.append("🔴 激光不活跃范围:")
        for i, inactive in enumerate(self.inactive_ranges):
            report.append(f"   • 范围{i+1}: {inactive}")
        report.append("")
        
        # 重叠分析
        overlap_analysis = self.analyze_overlaps()
        report.append("🔍 重叠分析:")
        report.append(f"   • 有重叠的不活跃范围: {overlap_analysis['summary']['ranges_with_overlaps']}")
        report.append(f"   • 无重叠的不活跃范围: {overlap_analysis['summary']['ranges_without_overlaps']}")
        report.append("")
        
        for overlap_data in overlap_analysis['overlaps']:
            inactive = overlap_data['inactive_range']
            report.append(f"   📌 不活跃范围 {inactive.start_time}s-{inactive.end_time}s 的重叠:")
            for overlap in overlap_data['overlaps']:
                segment = overlap['segment']
                duration = overlap['overlap_duration']
                report.append(f"      └─ 与 {segment.name} 重叠 {duration:.1f}秒")
        
        if overlap_analysis['no_overlaps']:
            report.append("   📌 无重叠的不活跃范围:")
            for inactive in overlap_analysis['no_overlaps']:
                report.append(f"      └─ {inactive.start_time}s-{inactive.end_time}s")
        report.append("")
        
        # 活动统计
        stats = self.calculate_laser_activity_stats()
        if stats:
            report.append("📈 激光活动统计:")
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
        
        # ASCII 时间线
        report.append("📊 时间线可视化:")
        ascii_timeline = self.generate_ascii_timeline()
        for line in ascii_timeline.split('\n'):
            report.append(f"   {line}")
        
        report.append("")
        report.append("=" * 60)
        
        return "\n".join(report)

def main():
    """主函数：分析用户提供的激光活动数据"""
    
    # 用户提供的原始数据
    raw_data = """
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
    
    print("🔍 激光点活动数据分析工具 (简化版)")
    print("=" * 50)
    
    # 创建分析器并解析数据
    analyzer = LaserActivityAnalyzer()
    parsed_data = analyzer.parse_chinese_data(raw_data)
    
    print(f"✅ 数据解析完成:")
    print(f"   • 发现 {len(parsed_data['segments'])} 个段落")
    print(f"   • 发现 {len(parsed_data['inactive_ranges'])} 个激光不活跃范围")
    print(f"   • 发现 {len(parsed_data['issues'])} 个数据问题")
    
    if parsed_data['issues']:
        print("\n⚠️ 发现的数据问题:")
        for issue in parsed_data['issues']:
            print(f"   • {issue}")
    
    # 修正时间范围问题
    print("\n🔧 修正时间范围问题...")
    fixed_count = analyzer.fix_time_ranges()
    
    # 生成详细报告
    print("\n📋 生成详细分析报告...")
    report = analyzer.generate_detailed_report()
    
    # 保存报告到文件
    report_file = "laser_analysis_report.txt"
    with open(report_file, 'w', encoding='utf-8') as f:
        f.write(report)
    
    print(f"✅ 报告已保存到: {report_file}")
    
    # 显示核心分析结果
    print("\n" + report)
    
    return analyzer

if __name__ == "__main__":
    analyzer = main()