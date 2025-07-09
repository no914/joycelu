import os
import re
import tempfile
from functools import lru_cache
from PIL import Image, ImageDraw, ImageFont
from typing import List, Dict
import numpy as np
from moviepy.editor import *
from moviepy.video.fx.all import speedx
from gtts import gTTS
from pptx import Presentation
import comtypes.client
from googletrans import Translator
import warnings
import time
from tqdm import tqdm
import psutil
import gc
import glob
import subprocess
import cv2

warnings.filterwarnings("ignore")

def translate_text(text, target_lang='en'):
    try:
        return Translator().translate(text, dest=target_lang).text
    except:
        return text
    
# 清理系统资源
gc.collect()

# 终止可能残留的ffmpeg进程
for proc in psutil.process_iter():
    if 'ffmpeg' in proc.name().lower():
        try:
            proc.kill()
        except:
            pass

class Config:
    OUTPUT_DIR = "output"
    SLIDE_IMAGE_DIR = "slides"
    LASER_COLOR = (255, 0, 0, 255)  # 带透明度的红色激光点
    LASER_RADIUS_RATIO = 0.04       # 点半径
    FADE_DURATION = 0.3             # 淡入淡出时间(秒)
    GLOW_COLOR = (255, 100, 100, 100)  # 光晕效果

class PPTSyncedConverter:
    def __init__(self):
        self.subtitle_style = {
            'font_size_ratio': 0.035,
            'bg_color': (0, 0, 128, 180),  # 半透明深蓝色背景
            'text_color': (255, 255, 255), # 白色文字
            'position_y': 0.85,            # 字幕垂直位置（0=顶部，1=底部）
            'padding': 20,                 # 文字与背景边距
            'max_width_ratio': 0.8        # 字幕最大宽度占比
        }
        self.progress_callback = None
        self._active_resources = []
        self.prs = None
        self._lock_files = set()
        self._check_dependencies()
        self._patch_moviepy_compatibility()
    
    def _patch_moviepy_compatibility(self):
        """修补moviepy兼容性问题"""
        try:
            from moviepy.audio.io.AudioFileClip import AudioFileClip
            from moviepy.video.io.VideoFileClip import VideoFileClip
            
            # 修补AudioFileClip
            def safe_close(self):
                if hasattr(self, 'reader') and self.reader:
                    try:
                        if hasattr(self.reader, 'close_proc'):
                            self.reader.close_proc()
                        if hasattr(self.reader, 'proc') and self.reader.proc:
                            self.reader.proc.terminate()
                    except:
                        pass
                if hasattr(self, 'close'):
                    super(AudioFileClip, self).close()
                    
            AudioFileClip.close = safe_close
            
            # 修补VideoFileClip
            def safe_write(self, *args, **kwargs):
                kwargs.update({
                    'ffmpeg_params': [
                        '-max_muxing_queue_size', '1024',
                        '-threads', '1'
                    ],
                    'preset': 'ultrafast',
                    'audio_codec': 'aac',
                    'temp_audiofile': 'temp_audio.mp3'
                })
                return super(VideoFileClip, self).write_videofile(*args, **kwargs)
                
            VideoFileClip.write_videofile = safe_write
            
        except Exception as e:
            print(f"兼容性修补失败: {str(e)}")
        
    def _check_dependencies(self):
        """检查依赖库版本"""
        import moviepy
        from pkg_resources import parse_version
        
        if parse_version(moviepy.__version__) < parse_version('1.0.0'):
            print("警告: moviepy版本过低，建议升级到1.0.0+")
            print("执行: pip install --upgrade moviepy")

    def _register_temp_file(self, path):
        """注册临时文件以便后续清理"""
        self._lock_files.add(os.path.abspath(path))

    def _cleanup_temp_files(self):
        """More robust temp file cleanup"""
        max_attempts = 3
        for file_path in list(self._lock_files):
            attempts = 0
            while attempts < max_attempts:
                try:
                    if os.path.exists(file_path):
                        os.unlink(file_path)
                    self._lock_files.remove(file_path)
                    break
                except Exception as e:
                    attempts += 1
                    if attempts >= max_attempts:
                        print(f"Failed to delete temp file {file_path}: {str(e)}")
                    time.sleep(0.5)

    def _cleanup_resources(self):
        """清理所有打开的媒体资源"""
        # 先关闭所有资源
        for res in self._active_resources[:]:
            try:
                if hasattr(res, 'close'):
                    res.close()
                self._active_resources.remove(res)
            except Exception as e:
                print(f"资源释放失败: {str(e)}")
        
        # 再清理临时文件
        self._cleanup_temp_files()
        
        # 强制终止可能残留的ffmpeg进程
        try:
            for proc in psutil.process_iter(['pid', 'name']):
                try:
                    if 'ffmpeg' in proc.info['name'].lower():
                        proc.kill()
                except:
                    pass
        except:
            pass
        
        # 强制垃圾回收
        try:
            import gc
            gc.collect()
        except:
            pass

    def __del__(self):
        """安全的析构函数"""
        try:
            self._cleanup_resources()
        except:
            pass  # 避免在解释器关闭时抛出异常

    def _ensure_path(self, path):
        """确保目录存在且路径有效"""
        if not path:
            raise ValueError("路径不能为空")
        
        dirname = os.path.dirname(path)
        if dirname:  # 如果不是当前目录
            os.makedirs(dirname, exist_ok=True)
        return os.path.abspath(path)

    def set_progress_callback(self, callback):
        self.progress_callback = callback

    def update_progress(self, message, progress=None):
        if self.progress_callback:
            self.progress_callback(message, progress)
        else:
            print(f"[{time.strftime('%H:%M:%S')}] {message}")

    def _get_slide_notes(self, slide):
        """安全获取幻灯片备注"""
        try:
            if not slide.has_notes_slide:
                return ""
            
            notes_slide = slide.notes_slide
            text = notes_slide.notes_text_frame.text
            print(f"✅ 提取到的备注内容: '{text}'")  # 调试输出
                
            return text
        except Exception as e:
            print(f"获取备注时出错: {str(e)}")
            return ""

    def ppt_to_images(self, pptx_path, output_dir, resolution=(1920, 1080)):
        powerpoint = None
        try:
            # 验证输入路径
            pptx_path = self._ensure_path(pptx_path)
            output_dir = self._ensure_path(output_dir)
            
            self.update_progress("正在初始化PPT")
            comtypes.CoInitialize()
            powerpoint = comtypes.client.CreateObject("PowerPoint.Application")
            deck = powerpoint.Presentations.Open(pptx_path)

            slide_count = deck.Slides.Count
            self.update_progress(f"开始转换PPT，共{slide_count}页...")
            
            for i in range(1, slide_count + 1):
                # 确保输出路径有效
                img_path = self._ensure_path(os.path.join(output_dir, f"slide_{i:03d}.png"))
                
                self.update_progress(f"正在转换第{i}/{slide_count}页...", i/slide_count)
                deck.Slides.Item(i).Export(img_path, "PNG")
                
                # 处理图片
                img = Image.open(img_path).convert("RGBA")
                if resolution != (1920, 1080):
                    img = img.resize(resolution, Image.LANCZOS)
                img.save(img_path)
            
            self.update_progress("PPT转换完成!")
            return True
        except Exception as e:
            print(f"PPT转换错误: {str(e)}")
            return False
        finally:
            if powerpoint:
                deck.Close()
                powerpoint.Quit()
                comtypes.CoUninitialize()

    def ensure_rgba(self, image_array):
        """Ensure image is RGBA with proper dimensions"""
        try:
            if image_array is None:
                return np.zeros((100, 100, 4), dtype=np.uint8)
                
            if len(image_array.shape) == 2:  # Grayscale
                return np.dstack([image_array]*3 + [np.ones(image_array.shape, dtype=np.uint8)*255])
            elif image_array.shape[2] == 3:  # RGB
                return np.dstack((image_array, np.ones(image_array.shape[:2], dtype=np.uint8)*255))
            elif image_array.shape[2] == 4:  # RGBA
                return image_array
            else:
                return np.zeros((100, 100, 4), dtype=np.uint8)
        except Exception as e:
            print(f"Image conversion error: {str(e)}")
            return np.zeros((100, 100, 4), dtype=np.uint8)

    def is_english(self, text):
        return all(ord(c) < 128 for c in text)
    
    def calculate_font_size(self, text, resolution):
        base_size = int(resolution[1] * self.subtitle_style['font_size_ratio'])
        max_size = int(resolution[1] * 0.05)
        min_size = int(resolution[1] * 0.025)
        
        if self.is_english(text) and len(text) > 30:
            return max(min_size, base_size - 2)
        return min(max_size, base_size)
    
    def wrap_text(self, text, font, max_width):
        max_pixel_width = max_width * self.subtitle_style['max_width_ratio']
        
        if not self.is_english(text):
            chars = list(text)
            lines = []
            current_line = ""
            for char in chars:
                test_line = current_line + char
                if font.getlength(test_line) <= max_pixel_width:
                    current_line = test_line
                else:
                    lines.append(current_line)
                    current_line = char
            if current_line:
                lines.append(current_line)
            return '\n'.join(lines)
        
        words = text.split()
        lines = []
        current_line = words[0] if words else ""
        
        for word in words[1:]:
            test_line = f"{current_line} {word}"
            if font.getlength(test_line) <= max_pixel_width:
                current_line = test_line
            else:
                lines.append(current_line)
                current_line = word
        if current_line:
            lines.append(current_line)
        
        final_lines = []
        for line in lines:
            if font.getlength(line) > max_pixel_width:
                split_line = ""
                for char in line:
                    if font.getlength(split_line + char) > max_pixel_width:
                        final_lines.append(split_line)
                        split_line = char
                    else:
                        split_line += char
                if split_line:
                    final_lines.append(split_line)
            else:
                final_lines.append(line)
        
        return '\n'.join(final_lines)
    
    def process_audio_segment(self, text_segment, lang, speed, temp_dir, index):
        # Use MP3 extension consistently
        try:
            if not text_segment or not text_segment.strip():
                duration = 1.0
                return None, duration
            
            clean_text = text_segment.strip()
            speed = max(0.5, min(3.0, float(speed))) if speed else 1.0

            estimated_duration = max(1.0, len(clean_text) * 0.15)
            estimated_duration = min(15.0, estimated_duration)

            print(f"尝试生成音频：文本长度={len(clean_text)}，预估时长={estimated_duration:.1f}秒")
            temp_audio_path = os.path.join(temp_dir, f"audio_{index}.mp3")

            # 1. Generate TTS audio
            try:
                tts = gTTS(text=clean_text, lang=lang, slow=False)
                tts.save(temp_audio_path)

                if not os.path.exists(temp_audio_path) or os.path.getsize(temp_audio_path) < 1024:
                    raise ValueError("TTS生成的文件无效或太小")
                
                try:
                    audio = AudioFileClip(temp_audio_path)
                    if audio.duration <= 0 or audio.duration > 30:
                        audio.close()
                        raise ValueError(f"音频时常异常: {audio.duration}")

                    # 5. Speed adjustment
                    if speed != 1.0:
                        try:
                            audio = audio.fx(speedx, speed)
                        except Exception as e:
                            print(f"语速调整失败，使用原速: {str(e)}")

                    final_duration = audio.duration
                    if final_duration <= 0 or final_duration > 30:
                        audio.close()
                        raise ValueError(f"最终音频时长异常：{final_duration}")
                    print(f"音频生成成功：时长={final_duration:.1f}秒")
                    return audio, final_duration

                except Exception as e:
                    print(f"音频加载失败: {str(e)}")
                    return None, estimated_duration
            except Exception as e:
                print(f"TTS生成失败：{str(e)}")
                return None, estimated_duration
        except Exception as e:
            print(f"音频处理完全失败: {str(0)}")
            fallback_duration = 2.0
            return None, fallback_duration
        
        finally:
            # Clean up temporary files
            temp_files = [
                os.path.join(temp_dir, f"audio_{index}.mp3"),
                os.path.join(temp_dir, f"final_audio_{index}.mp3")
            ]
            for f in temp_files:
                if os.path.exists(f):
                    try:
                        os.unlink(f)
                    except:
                        pass
    
    def create_synced_slide(self, img_path, slide_index, lang, speed, temp_dir, index, resolution):
        """创建带字幕和激光笔动画的幻灯片视频（完整安全版本）"""
        try:
            # ==================== 1. 初始化检查 ====================
            gc.collect()
            if psutil.virtual_memory().percent > 85:
                time.sleep(5)

            # 验证输入参数
            if not os.path.exists(img_path):
                raise FileNotFoundError(f"幻灯片图片不存在: {img_path}")
            
            width, height = resolution
            if width <= 0 or height <= 0:
                raise ValueError(f"无效的分辨率: {width}x{height}")

            # ==================== 2. 加载幻灯片内容 ====================
            slide = self.prs.slides[slide_index]
            notes_text = self._get_slide_notes(slide) or ""
            
            # 提取纯净文本（移除激光指令）
            display_text = re.sub(r'\[cursor:\s*\d+,\s*\d+\]|\[cursor:\s*off\]', '', notes_text).strip()
            
            # 解析激光点（带安全坐标检查）
            laser_points = self.parse_laser_actions(notes_text, None, width, height, lang, speed)

            # ==================== 3. 文本预处理 ====================
            # 自动翻译非目标语言文本
            if display_text and lang != 'zh-cn' and any('\u4e00' <= c <= '\u9fff' for c in display_text):
                display_text = translate_text(display_text, 'en')
            
            # 分割文本为适合语音合成的段落
            segments = self._split_text_segments(notes_text)
            if not segments:
                segments = [display_text] if display_text else [""]

            # ==================== 4. 生成语音和时长数据 ====================
            video_clips = []
            segment_data = []
            current_time = 0.0

            for i, segment in enumerate(segments):
                try:
                    clean_segment = re.sub(r'\[cursor:\s*\d+,\s*\d+\]|\[cursor:\s*off\]', '', segment).strip()

                    audio = None
                    duration = 2.0
                    
                    # 生成语音（带错误处理）
                    if clean_segment:
                        try:
                            audio, duration = self.process_audio_segment(
                                clean_segment, lang, speed, temp_dir, f"{index}_{i}"
                            )
                            print(f"段落{i}: 音频={'有效' if audio else '无'}，时长={duration:.1f}秒")
                        except Exception as e:
                            print(f"段落{i} 音频生成异常：{str(e)}")
                            audio = None
                            duration = max(1.0, len(clean_segment) * 0.15)
                    else:
                        duration = 3.0 if slide_index == 0 else 1.5  # 默认时长
                        print(f"段落{i}: 空文本，使用默认时长{duration:.1f}秒")

                    duration = max(0.5, min(15.0, duration))

                    segment_data.append({
                        "text": segment,
                        "clean_text": clean_segment,
                        "audio": audio,
                        "duration": duration,
                        "start_time": current_time,
                        "end_time": current_time + duration
                    })
                    current_time += duration

                except Exception as e:
                    print(f"段落{i}处理失败: {str(e)}")
                    fallback_duration = 2.0
                    segment_data.append({
                        "text": segment if 'segment' in locals() else "fallback",
                        "clean_text": "",
                        "audio": None,
                        "duration": fallback_duration,
                        "start_time": current_time,
                        "end_time": current_time + fallback_duration
                    })
                    current_time += fallback_duration

            # ==================== 5. 加载背景图片（只加载一次） ====================
            try:
                bg_img = Image.open(img_path).convert('RGBA')
                if bg_img.size != (width, height):
                    bg_img = bg_img.resize((width, height), Image.LANCZOS)
            except Exception as e:
                print(f"背景图片处理失败: {str(e)}")
                bg_img = Image.new('RGBA', (width, height), (255, 255, 255, 255))

            # ==================== 6. 生成视频片段 ====================
            for seg_idx, seg in enumerate(segment_data):
                try:
                    if not seg or "duration" not in seg or seg['duration'] <= 0:
                        print(f"跳过无效的段落{seg_idx}")
                        continue
                    seg['duration'] = min(15.0, max(0.5, seg['duration']))

                    # 创建透明覆盖层
                    overlay = Image.new("RGBA", (width, height), (0, 0, 0, 0))
                    draw = ImageDraw.Draw(overlay)

                    # ==================== 7. 安全生成字幕 ====================
                    if seg.get("clean_text"):
                        try:
                            # 生成安全字幕图像
                            subtitle_img = self._generate_safe_subtitle(
                                text=seg["clean_text"],
                                bg_img=bg_img,
                                max_width=width,
                                max_height=height
                            )
                            
                            if subtitle_img and subtitle_img.size[0] > 0 and subtitle_img.size[1] > 0:
                                # 计算安全位置（带边界检查）
                                y_pos = int(height * self.subtitle_style['position_y'] - subtitle_img.height // 2)
                                y_pos = max(10, min(height - subtitle_img.height - 10, y_pos))
                                
                                x_pos = (width - subtitle_img.width) // 2
                                x_pos = max(10, min(width - subtitle_img.width - 10, x_pos))
                                
                                # 合并字幕到覆盖层
                                overlay.paste(subtitle_img, (x_pos, y_pos), subtitle_img)
                        except Exception as e:
                            print(f"字幕生成失败: {str(e)}")

                    # ==================== 8. 绘制激光笔动画 ====================
                    if laser_points:
                        try:
                            radius = int(height * Config.LASER_RADIUS_RATIO)
                            
                            # 找出当前时间段活跃的激光点
                            active_points = []
                            for p in laser_points: 
                                if (p.get('end', 0) > seg.get("start_time", 0) and
                                    p.get('start', 0) < seg.get("end_time", seg.get("duration", 0))):
                                    active_points.append(p)
                            
                            for p in active_points:
                                try:
                                    if 'x' not in p or 'y' not in p:
                                        continue

                                    # 安全坐标检查
                                    x = max(radius, min(width - radius, int(p['x'])))
                                    y = max(radius, min(height - radius, int(p['y'])))
                                        
                                    # 绘制光晕效果
                                    draw.ellipse(
                                        [
                                            max(0, x - radius - 10),
                                            max(0, y - radius - 10),
                                            min(width - 1, x + radius + 10),
                                            min(height - 1, y + radius + 10)
                                        ],
                                        fill=Config.GLOW_COLOR
                                    )
                                        
                                    # 绘制激光点核心
                                    draw.ellipse(
                                        [
                                            max(0, x - radius),
                                            max(0, y - radius),
                                            min(width - 1, x + radius),
                                            min(height - 1, y + radius)
                                        ],
                                        fill=Config.LASER_COLOR
                                    )
                                except Exception as e:
                                    print(f"激光点绘制失败: {str(e)}")
                                    continue
                        except Exception as e:
                            print(f"激光点处理失败: {str(e)}")

                    # ==================== 9. 合并图层 ====================
                    try:
                        final_img = Image.alpha_composite(bg_img, overlay)
                    except Exception as e:
                        print(f"图层合并失败: {str(e)}")
                        final_img = bg_img
                    
                    # 确保图像数组格式正确
                    try:
                        img_array = np.array(final_img)
                        if img_array.ndim != 3 or img_array.shape[2] not in [3, 4]:
                            img_array = np.zeros((height, width, 3), dtype = np.uint8)
                            img_array.fill(255)
                        elif img_array.shape[2] == 4:
                            img_array = img_array[:, :, :3]
                    except Exception as e:
                        print(f"图像数组转换失败: {str(e)}")
                        img_array = np.full((height, width, 3), 255, dtype = np.uint8)

                    # ==================== 10. 创建视频片段 ====================
                    try:
                        final_clip = ImageClip(img_array).set_duration(seg["duration"])

                        has_valid_audio = False
                        if seg.get("audio") and hasattr(seg["audio"], 'duration'):
                            try:
                                audio_duration = seg["audio"].duration
                                if audio_duration > 0 and abs(audio_duration - seg["duration"]) < 2.0: 
                                    final_clip = final_clip.set_audio(seg["audio"])
                            except Exception as e:
                                print(f"音频添加失败: {str(e)}")
                        
                        # 保存临时文件
                        temp_path = os.path.join(temp_dir, f"seg_{index}_{seg_idx}.mp4")
                        final_clip.write_videofile(
                            temp_path,
                            fps=15,
                            codec='libx264',
                            audio_codec='aac',
                            preset='ultrafast',
                            threads=1,
                            ffmpeg_params=['-pix_fmt', 'yuva420p', '-max_muxing_queue_size', '1024', '-avoid_negative_ts', 'make_zero'],
                            logger=None,
                            verbose = False
                        )
                        if os.path.exists(temp_path) and os.path.getsize(temp_path) > 1024:
                            video_clips.append(temp_path)
                        else:
                            print(f"生成的视频文件无效: {temp_path}")

                        final_clip.close()
                        if seg.get("audio"):
                            try:
                                seg["audio"].close()
                            except:
                                pass
                    except Exception as e:
                        print(f"视频片段创建失败: {str(e)}")
                        continue
                    
                except Exception as e:
                    print(f"创建幻灯片片段时出错: {str(e)}")
                    continue

            # ==================== 11. 合并所有片段 ====================
            if video_clips:
                try:
                    loaded_clips = [VideoFileClip(v) for v in video_clips]
                    return concatenate_videoclips(loaded_clips)
                except Exception as e:
                    print(f"合并片段失败: {str(e)}")
            
            # ==================== 12. 回退方案 ====================
            return self._create_fallback_clip(img_path, temp_dir, index)

        except Exception as e:
            print(f"创建幻灯片时出错: {str(e)}")
            return self._create_fallback_clip(img_path, temp_dir, index)

    def _generate_safe_subtitle(self, text, bg_img, max_width, max_height):
        try:
            # Get image dimensions from bg_img without using the background
            if isinstance(bg_img, Image.Image):
                img_size = bg_img.size
            else:
                # Assume it's a file path
                temp_img = Image.open(bg_img).convert("RGBA")
                img_size = temp_img.size
                
            # Create ONLY the subtitle overlay (no background)
            overlay = Image.new("RGBA", img_size, (0, 0, 0, 0))
            draw = ImageDraw.Draw(overlay)
            
            font_size = self.calculate_font_size(text, (max_width, max_height))
            
            try:
                if self.is_english(text):
                    font = ImageFont.truetype("arial.ttf", font_size)
                else:
                    font = ImageFont.truetype("simhei.ttf", font_size)
            except:
                font = ImageFont.load_default(size=font_size)
                print("警告：使用默认字体，可能不支持中文")

            wrapped_text = self.wrap_text(text, font, max_width)
            lines = wrapped_text.split('\n')
            
            line_bboxes = [font.getbbox(line) for line in lines]
            line_widths = [bbox[2] - bbox[0] for bbox in line_bboxes]
            line_heights = [bbox[3] - bbox[1] for bbox in line_bboxes]
            total_height = sum(line_heights)
            max_line_width = max(line_widths)
            
            bg_width = max_line_width + 2 * self.subtitle_style['padding']
            bg_height = total_height + 2 * self.subtitle_style['padding']
            y_position = img_size[1] * self.subtitle_style['position_y']
            bg_y1 = y_position - bg_height // 2
            bg_y2 = y_position + bg_height // 2
            
            draw.rectangle(
                [(img_size[0] - bg_width) // 2, bg_y1,
                (img_size[0] + bg_width) // 2, bg_y2],
                fill=self.subtitle_style['bg_color']
            )
            
            current_y = bg_y1 + self.subtitle_style['padding']
            for i, line in enumerate(lines):
                text_width = font.getlength(line)
                x = (img_size[0] - text_width) // 2
                draw.text(
                    (x, current_y),
                    line,
                    font=font,
                    fill=self.subtitle_style['text_color'],
                    stroke_width=2,
                    stroke_fill=(0, 0, 0)
                )
                current_y += line_heights[i]
            
            # Return ONLY the subtitle overlay, not composited with background
            return overlay
        
        except Exception as e:
            print(f"生成字幕图片失败: {str(e)}")
            # Create error overlay (transparent background with error text)
            error_overlay = Image.new("RGBA", (max_width, max_height), (0, 0, 0, 0))
            draw = ImageDraw.Draw(error_overlay)
            # Add semi-transparent red background for error
            draw.rectangle([(10, 10), (max_width-10, 50)], fill=(255, 0, 0, 128))
            draw.text((15, 15), "Subtitle Error", fill=(255, 255, 255, 255))
            return error_overlay

    def _split_text_lines(self, text, max_chars, max_pixel_width):
        """将文本分割为适合显示的行"""
        lines = []
        
        if self.is_english(text):
            words = text.split()
            current_line = ""
            for word in words:
                if len(current_line + " " + word) <= max_chars:
                    current_line += (" " + word if current_line else word)
                else:
                    lines.append(current_line)
                    current_line = word
            if current_line:
                lines.append(current_line)
        else:  # 中文处理
            current_line = ""
            for char in text:
                if len(current_line + char) <= max_chars:
                    current_line += char
                else:
                    lines.append(current_line)
                    current_line = char
            if current_line:
                lines.append(current_line)
        
        return lines

    def _calculate_optimal_font_size(self, lines, max_width, max_height):
        """动态计算最佳字体大小"""
        max_font = int(max_height * 0.05)
        min_font = 12
        
        for size in range(max_font, min_font, -1):
            try:
                font = ImageFont.truetype("simhei.ttf", size)
            except:
                font = ImageFont.load_default(size=size)
            
            # 检查所有行是否适应
            fits = True
            for line in lines:
                if font.getbbox(line)[2] > max_width * 0.9:
                    fits = False
                    break
                    
            if fits and (len(lines) * (font.getbbox("A")[3] + 5)) < max_height * 0.3:
                return size
        
        return min_font

    def _create_fallback_clip(self, img_path, temp_dir, index):
        """创建后备视频片段（当主流程失败时使用）"""
        try:
            bg_img = Image.open(img_path).convert('RGB')
            bg_clip = ImageClip(np.array(bg_img)).set_duration(3.0)
            temp_path = os.path.join(temp_dir, f"seg_{index}_fallback.mp4")
            bg_clip.write_videofile(temp_path, fps=15)
            return VideoFileClip(temp_path)
        except:
            # 终极后备方案
            blank = np.zeros((100, 100, 3), dtype=np.uint8)
            return ImageClip(blank).set_duration(3.0)
        
    

    def _split_text_segments(self, text):
        """辅助方法：分割文本为语音段落"""
        segments = []
        current_segment = ""
        cursor_pattern = re.compile(r'(\[cursor:\s*\d+,\s*\d+\]|\[cursor:\s*off\])')
        
        for line in text.split('\n'):
            if line.strip():
                parts = cursor_pattern.split(line)
                for part in parts:
                    if not part:
                        continue
                        
                    if cursor_pattern.match(part):
                        current_segment += part
                        if current_segment.strip():
                            segments.append(current_segment)
                            current_segment = ""
                    else:
                        current_segment += part
                        if part.endswith(('。', '!', '?', ';')):
                            segments.append(current_segment)
                            current_segment = ""
        
        if current_segment:
            segments.append(current_segment)
        
        return segments

    def estimate_segment_duration(self, text, lang, speed):
        """估算语音段落持续时间"""
        # 基本估算：每分钟约150个单词（英文）或300个汉字（中文）
        try:
            if not text or not isinstance(text, str):
                return 2.0
            clean_text = re.sub(r'\[cursor:\s*\d+,\s*\d+\]|\[cursor:\s*off\]', '', text).strip()
            if not clean_text:
                return 1.5
            speed = max(0.5, min(3.0, float(speed))) if speed else 1.0
            if lang and lang.startwith('zh'):
                char_count = len(clean_text)
                chars_per_minute = 240
                duration = max(1.0, (char_count / chars_per_minute) * 60 / speed)
            else:
                word_count = len(clean_text.split())
                words_per_minute = 120
                duration = max(1.0, (word_count / words_per_minute) * 60 / speed)

            min_duration = max(1.0, len(clean_text) * 0.1)
            duration = max(min_duration, duration)
            return min(duration, 15.0)  # 限制最长10秒
        except Exception as e:
            print(f"持续时间估算失败: {str(e)}")
            return 3.0

    def create_laser_animation(self, slide_path, points, resolution):
        try:
            width, height = resolution
            radius = int(height * Config.LASER_RADIUS_RATIO)
            
            # 创建调试目录
            debug_dir = os.path.abspath("laser_debug")
            os.makedirs(debug_dir, exist_ok=True)
            
            def make_frame(t):
                # 创建透明图层
                img = Image.new("RGBA", (width, height), (0, 0, 0, 0))
                draw = ImageDraw.Draw(img)
                
                # 找出当前时间点活跃的激光点
                active_points = [p for p in points if p['start'] <= t <= p['end']]
                
                for p in active_points:
                    x, y = int(p['x']), int(p['y'])
                    print(f"在位置 ({x}, {y}) 绘制激光点")  # 调试输出
                    
                    # 绘制光晕效果
                    draw.ellipse(
                        [x-radius-10, y-radius-10, x+radius+10, y+radius+10],
                        fill=Config.GLOW_COLOR
                    )
                    
                    # 绘制激光点核心
                    draw.ellipse(
                        [x-radius, y-radius, x+radius, y+radius],
                        fill=Config.LASER_COLOR
                    )
                
                # 保存调试图片
                debug_path = os.path.join(debug_dir, f"laser_{int(t*100)}.png")
                img.save(debug_path)
                
                return np.array(img)
            
            duration = max(p['end'] for p in points) if points else 1.0
            return VideoClip(make_frame, duration=duration)
        
        except Exception as e:
            print(f"激光动画创建失败: {str(e)}")
            return None

    def adjust_laser_timing(self, points, new_duration):
        if not points:
            return points
        
        original_duration = points[-1]["end"]
        if original_duration <= 0:
            return points
        
        ratio = new_duration / original_duration
        return [
            {
                "x": p["x"],
                "y": p["y"],
                "start": p["start"] * ratio,
                "end": p["end"] * ratio
            }
            for p in points
        ]
        
    def split_text_for_speech(self, text, lang):
        if lang.startswith('zh'):
            sentences = re.split('([，。！？])', text)
            return [s for s in [''.join(x) for x in zip(sentences[::2], sentences[1::2])] if s.strip()]
        else:
            return [s.strip() for s in re.split(r'[.!?]', text) if s.strip()]
        
    def calculate_word_timing(self, text: str, lang: str = 'zh-cn', speed: float = 1.0) -> Dict:
        """计算每个单词在语音中的时间位置"""
        clean_text = re.sub(r'\[cursor:\s*\d+,\s*\d+\]|\[cursor:\s*off\]', '', text).strip()
        
        if not clean_text:
            return {}
        
        # 基础语音参数
        speed = max(0.5, min(3.0, float(speed))) if speed else 1.0
        
        if lang and lang.startswith('zh'):
            # 中文：按字符分割
            chars = list(clean_text)
            chars_per_second = 3.5 * speed  # 中文每秒约3.5个字符
            
            word_timings = {}
            current_time = 0.0
            
            for i, char in enumerate(chars):
                if char.strip():  # 跳过空格
                    char_duration = 1.0 / chars_per_second
                    word_timings[i] = {
                        'word': char,
                        'start': current_time,
                        'end': current_time + char_duration
                    }
                    current_time += char_duration
                else:
                    current_time += 0.1  # 空格/标点的停顿时间
        else:
            # 英文：按单词分割
            words = clean_text.split()
            words_per_second = 2.5 * speed  # 英文每秒约2.5个单词
            
            word_timings = {}
            current_time = 0.0
            
            for i, word in enumerate(words):
                word_duration = len(word) / (words_per_second * 3)  # 根据单词长度调整
                word_duration = max(0.2, min(1.0, word_duration))  # 限制单词时长
                
                word_timings[i] = {
                    'word': word,
                    'start': current_time,
                    'end': current_time + word_duration
                }
                current_time += word_duration + 0.1  # 单词间停顿
        
        return word_timings

    def parse_laser_actions(self, note_text: str, duration: float = None, 
                        slide_width: int = 1920, slide_height: int = 1080,
                        lang: str = 'zh-cn', speed: float = 1.0) -> List[Dict]:
        """解析激光笔指令并计算基于单词的精确时间位置"""
        # 输入验证
        if not note_text or not isinstance(note_text, str):
            return []
        
        if slide_width <= 0 or slide_height <= 0:
            print(f"警告: 无效的分辨率 {slide_width}x{slide_height}，使用默认值")
            slide_width, slide_height = 1280, 720

        print(f"🔍 原始备注文本: '{note_text}'")
        
        try:
            # 1. 获取单词时间映射
            word_timings = self.calculate_word_timing(note_text, lang, speed)
            
            # 2. 找到激光指令的位置和前后文本
            cursor_pattern = re.compile(r'\[cursor:\s*(\d+),\s*(\d+)\]|\[cursor:\s*off\]')
            laser_points = []
            
            # 3. 分析每个激光指令的上下文
            for match in cursor_pattern.finditer(note_text):
                cursor_cmd = match.group()
                cursor_pos = match.start()
                
                # 找到指令前的文本
                text_before = note_text[:cursor_pos]
                # 移除之前的激光指令，只保留纯文本
                clean_text_before = re.sub(r'\[cursor:\s*\d+,\s*\d+\]|\[cursor:\s*off\]', '', text_before)
                
                if 'off' in cursor_cmd.lower():
                    # 激光点结束：找到指令前最后一个单词的结束时间
                    trigger_time = self._find_word_end_time(clean_text_before, word_timings, lang)
                    
                    # 更新最后一个激光点的结束时间
                    if laser_points:
                        laser_points[-1]['end'] = trigger_time + 0.1  # 稍微延迟一点
                        print(f"🔴 激光点结束时间: {trigger_time:.2f}秒 (在文本 '{clean_text_before}' 之后)")
                else:
                    # 激光点开始：解析坐标并设置开始时间
                    coords = re.findall(r'\d+', cursor_cmd)
                    if len(coords) >= 2:
                        try:
                            # 坐标转换
                            x_percent = min(100, max(0, float(coords[0])))
                            y_percent = min(100, max(0, float(coords[1])))
                            x = max(0, min(slide_width - 1, int(slide_width * x_percent / 100)))
                            y = max(0, min(slide_height - 1, int(slide_height * y_percent / 100)))
                            
                            # 计算触发时间：指令前最后一个单词的结束时间
                            trigger_time = self._find_word_end_time(clean_text_before, word_timings, lang)
                            
                            laser_point = {
                                'x': x,
                                'y': y,
                                'start': trigger_time,
                                'end': trigger_time + 2.0  # 默认持续2秒，如果没有off指令
                            }
                            laser_points.append(laser_point)
                            print(f"🟢 激光点开始时间: {trigger_time:.2f}秒 (在文本 '{clean_text_before}' 之后)")
                            
                        except (ValueError, TypeError) as e:
                            print(f"坐标解析错误: {str(e)}")
                            continue
            
            # 4. 验证和调整时间
            validated_points = []
            total_duration = max(word_timings[k]['end'] for k in word_timings) if word_timings else 3.0
            
            for point in laser_points:
                start = max(0.0, min(total_duration - 0.1, point['start']))
                end = max(start + 0.1, min(total_duration, point['end']))
                
                if end > start and start >= 0:
                    validated_points.append({
                        'x': point['x'],
                        'y': point['y'],
                        'start': start,
                        'end': end
                    })
            
            print(f"🔧 解析后的激光点: {validated_points}")
            return validated_points
            
        except Exception as e:
            print(f"激光点解析错误: {str(e)}")
            return []

    def _find_word_end_time(self, text_before: str, word_timings: Dict, lang: str) -> float:
        """找到指定文本中最后一个单词的结束时间"""
        if not text_before.strip() or not word_timings:
            return 0.0
        
        clean_text = text_before.strip()
        
        if lang and lang.startswith('zh'):
            # 中文：查找最后一个字符的时间
            char_count = len(clean_text)
            if char_count > 0:
                # 在word_timings中找到对应位置的字符
                for i in range(char_count - 1, -1, -1):
                    if i in word_timings:
                        return word_timings[i]['end']
        else:
            # 英文：查找最后一个单词的时间
            words = clean_text.split()
            if words:
                word_count = len(words)
                if word_count > 0:
                    # 在word_timings中找到对应位置的单词
                    for i in range(word_count - 1, -1, -1):
                        if i in word_timings:
                            return word_timings[i]['end']
        
        return 0.0

    def make_layers_compatible(self, layers):
        """确保所有图层通道一致"""
        compatible_layers = []
        
        for layer in layers:
            if layer is None:
                continue
                
            try:
                frame = layer.get_frame(0) if hasattr(layer, 'get_frame') else layer.img
                frame = self.ensure_rgba(frame)
                
                # 重建兼容的clip
                if hasattr(layer, 'get_frame'):
                    new_clip = VideoClip(lambda t: frame, duration=layer.duration)
                else:
                    new_clip = ImageClip(frame)
                
                compatible_layers.append(new_clip.set_duration(layer.duration))
                
            except Exception as e:
                print(f"图层转换失败: {str(e)}")
                continue
                
        return compatible_layers

    def convert_ppt_to_video(self, pptx_path, output_path, lang='zh-cn', resolution=(854, 480), speed=1.0, fps=15):
        try:
            # ==================== 1. 初始化验证 ====================
            self.update_progress("正在验证输入文件...")
            pptx_path = self._ensure_path(pptx_path)
            output_path = self._ensure_path(output_path)
            
            if not os.path.exists(pptx_path):
                raise FileNotFoundError(f"PPTX文件不存在: {pptx_path}")
                
            if not pptx_path.lower().endswith('.pptx'):
                raise ValueError("仅支持.pptx格式的文件")

            # ==================== 2. 加载PPTX ====================
            self.update_progress("正在加载PPTX文件...")
            try:
                self.prs = Presentation(pptx_path)
                if len(self.prs.slides) == 0:
                    raise ValueError("PPTX中没有幻灯片")
            except Exception as e:
                raise ValueError(f"PPTX加载失败: {str(e)}")

            # ==================== 3. 创建临时工作区 ====================
            temp_dir = tempfile.mkdtemp(prefix='ppt2video_')
            try:
                temp_dir = self._ensure_path(temp_dir)
                self.update_progress(f"临时工作区: {temp_dir}")
                
                # ==================== 4. PPT转图片 ====================
                slide_images = []
                self.update_progress("正在将PPT转换为图片...")
                if not self.ppt_to_images(pptx_path, temp_dir, resolution):
                    raise RuntimeError("PPT转图片失败")
                
                slide_images = sorted(glob.glob(os.path.join(temp_dir, "slide_*.png")))
                if len(slide_images) != len(self.prs.slides):
                    raise RuntimeError(f"生成的图片数量({len(slide_images)})与幻灯片数量({len(self.prs.slides)})不匹配")

                # ==================== 5. 逐页处理 ====================
                video_segments = []
                success_count = 0
                
                for i, img_path in enumerate(slide_images):
                    try:
                        # 内存监控
                        mem = psutil.virtual_memory()
                        if mem.available < 200 * 1024 * 1024:
                            self.update_progress("内存不足，暂停处理...")
                            time.sleep(5)
                            gc.collect()
                        
                        # 处理单页
                        self.update_progress(f"正在处理第{i+1}/{len(slide_images)}页...", (i+1)/len(slide_images))
                        
                        result = self.create_synced_slide(
                            img_path=img_path,
                            slide_index=i,
                            lang=lang,
                            speed=speed,
                            temp_dir=temp_dir,
                            index=i,
                            resolution=resolution
                        )
                        
                        if result:
                            temp_video = os.path.join(temp_dir, f"clip_{i}.mp4")
                            result.write_videofile(temp_video, fps=fps)
                            if os.path.exists(temp_video) and os.path.getsize(temp_video) > 0:
                                video_segments.append(temp_video)
                                success_count += 1
                            else:
                                self.update_progress(f"生成的第{i+1}页视频无效")
                        else:
                            self.update_progress(f"跳过无效的第{i+1}页")
                            
                    except Exception as e:
                        self.update_progress(f"第{i+1}页处理失败: {str(e)}")
                        continue

                # ==================== 6. 合并视频 ====================
                if success_count > 0:
                    self.update_progress("正在合并视频片段...")
                    
                    # 使用moviepy合并视频
                    clips = []
                    for seg in video_segments:
                        try:
                            clip = VideoFileClip(seg)
                            clips.append(clip)
                        except Exception as e:
                            self.update_progress(f"加载视频片段失败: {seg} - {str(e)}")
                    
                    if clips:
                        final_clip = concatenate_videoclips(clips)
                        final_clip.write_videofile(
                            output_path,
                            fps=fps,
                            threads=4,
                            preset='ultrafast',
                            audio_codec='aac',
                            verbose=False
                        )
                        final_clip.close()
                        
                        for clip in clips:
                            clip.close()
                        
                        return True
                    
                return False
                
            finally:
                # 清理临时目录
                try:
                    for root, dirs, files in os.walk(temp_dir, topdown=False):
                        for name in files:
                            try:
                                os.unlink(os.path.join(root, name))
                            except:
                                pass
                        for name in dirs:
                            try:
                                os.rmdir(os.path.join(root, name))
                            except:
                                pass
                    os.rmdir(temp_dir)
                except:
                    pass
                
        except Exception as e:
            self.update_progress(f"转换失败: {str(e)}")
            return False
        finally:
            self._cleanup_resources()

    def _check_system_resources(self):
        """检查系统资源状态"""
        mem = psutil.virtual_memory()
        if mem.percent > 90:
            self.update_progress(f"系统内存使用率高: {mem.percent}%")
            gc.collect()
            time.sleep(5)
            return False
        return True

def get_users_choice():
    print("请选择语音语言")
    print("1. 中文")
    print("2. 英文")
    while True:
        lang_choice = input("请输入选项(1-2): ").strip()
        if lang_choice in ('1', '2'):
            lang = 'zh-cn' if lang_choice == '1' else 'en'
            break
        print("无效输入，请重新选择")

    print("\n请选择分辨率")
    print("1. 1920x1080 (全高清)")
    print("2. 1280x720 (高清)")
    while True:
        resolution_choice = input("请输入选项(1-2): ").strip()
        if resolution_choice == '1':
            resolution = (1920, 1080)
            break
        elif resolution_choice == '2':
            resolution = (1280, 720)
            break
        else:
            print("无效输入，请重新选择")
        
    print("\n请选择语音速度")
    print("1. 慢速 (0.75x)")
    print("2. 正常 (1x)")
    print("3. 快速 (1.5x)")
    while True:
        speed_choice = input("请输入选项(1-3): ").strip()
        if speed_choice == '1':
            speed = 0.75
            break
        elif speed_choice == '2':
            speed = 1
            break
        elif speed_choice == '3':
            speed = 1.5
            break
        else:
            print("无效输入，请重新选择")

    return lang, resolution, speed

if __name__ == "__main__":

    converter = PPTSyncedConverter()
    converter.subtitle_style['bg_color'] = (0, 0, 128, 180)  # 设置半透明深蓝色背景

    lang, resolution, speed = get_users_choice()

    def progress_callback(message, progress=None):
        if progress is not None:
            print(f"[{time.strftime('%H:%M:%S')}] {message} ({progress*100:.1f}%)")
        else:
            print(f"[{time.strftime('%H:%M:%S')}] {message}")
    
    converter.set_progress_callback(progress_callback)
    
    success = converter.convert_ppt_to_video(
        pptx_path="test.pptx",
        output_path="output.mp4",
        lang=lang,
        resolution=resolution, 
        speed=speed
    )
    print("转换成功!" if success else "转换失败")
