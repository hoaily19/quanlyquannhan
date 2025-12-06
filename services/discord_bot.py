"""
Discord Bot Service - Gửi thông báo hoạt động lên Discord
"""

try:
    import discord
    from discord.ext import commands
    DISCORD_AVAILABLE = True
except ImportError:
    DISCORD_AVAILABLE = False
    discord = None
    commands = None

import asyncio
import logging
from typing import Optional
import sys
from pathlib import Path
import time
import queue

# Thêm path để import config
sys.path.insert(0, str(Path(__file__).parent.parent))

try:
    from config.discord_config import (
        DISCORD_BOT_TOKEN,
        DISCORD_CHANNEL_ID,
        ENABLE_DISCORD_NOTIFICATIONS,
        NOTIFY_ON_PERSONNEL_ADD,
        NOTIFY_ON_PERSONNEL_UPDATE,
        NOTIFY_ON_PERSONNEL_DELETE,
        NOTIFY_ON_EXPORT,
        ENABLE_REMOTE_CONTROL,
        ADMIN_USER_IDS
    )
except ImportError:
    # Nếu không có config, sử dụng giá trị mặc định
    DISCORD_BOT_TOKEN = None
    DISCORD_CHANNEL_ID = None
    ENABLE_DISCORD_NOTIFICATIONS = False
    NOTIFY_ON_PERSONNEL_ADD = False
    NOTIFY_ON_PERSONNEL_UPDATE = False
    NOTIFY_ON_PERSONNEL_DELETE = False
    NOTIFY_ON_EXPORT = False
    ENABLE_REMOTE_CONTROL = False
    ADMIN_USER_IDS = []

logger = logging.getLogger(__name__)


class DiscordBotService:
    """Service quản lý kết nối và gửi thông báo lên Discord"""
    
    def __init__(self):
        self.bot = None
        self.is_connected = False
        self.channel = None
        self.loop = None
        self.thread = None
        self.message_queue = queue.Queue()  # Queue để lưu thông báo chờ gửi
        self.connection_event = None  # Event để đợi bot kết nối
        
    async def _start_bot(self):
        """Khởi động bot Discord"""
        logger.info("=== BẮT ĐẦU KHỞI ĐỘNG DISCORD BOT ===")
        
        if not DISCORD_AVAILABLE:
            logger.error("❌ Thư viện discord.py chưa được cài đặt. Chạy: pip install discord.py")
            return
            
        if not DISCORD_BOT_TOKEN:
            logger.error("❌ Discord bot token không được cấu hình")
            return
            
        if not ENABLE_DISCORD_NOTIFICATIONS:
            logger.warning("⚠️ Thông báo Discord đã bị tắt")
            return
        
        logger.info(f"Token: {DISCORD_BOT_TOKEN[:20]}...")
        logger.info(f"Channel ID: {DISCORD_CHANNEL_ID}")
        
        try:
            # Cần bật message_content để bot có thể nhận lệnh từ Discord
            # LƯU Ý QUAN TRỌNG: 
            # 1. Bot CẦN "MESSAGE CONTENT INTENT" để nhận lệnh từ Discord
            # 2. Vào: https://discord.com/developers/applications/
            # 3. Chọn bot của bạn > Tab "Bot" 
            # 4. Tìm phần "Privileged Gateway Intents"
            # 5. BẬT "MESSAGE CONTENT INTENT" và LƯU
            # 6. Nếu không bật, bot sẽ không thể nhận lệnh !shutdown, !restart, v.v.
            
            intents = discord.Intents.default()
            # Tạm thời tắt để bot có thể kết nối, nhưng CẦN BẬT trong Developer Portal để nhận lệnh
            intents.message_content = False  # TẠM THỜI TẮT - CẦN BẬT TRONG DEVELOPER PORTAL!
            
            self.bot = commands.Bot(command_prefix='!', intents=intents)
            
            # Thêm commands để điều khiển ứng dụng
            if ENABLE_REMOTE_CONTROL:
                @self.bot.command(name='shutdown', aliases=['tắt', 'tat', 'off'])
                async def shutdown_command(ctx):
                    """Lệnh tắt ứng dụng"""
                    # Kiểm tra quyền
                    if ADMIN_USER_IDS and str(ctx.author.id) not in ADMIN_USER_IDS:
                        await ctx.send("❌ Bạn không có quyền sử dụng lệnh này!")
                        return
                    
                    logger.warning(f"⚠️ Lệnh tắt ứng dụng từ Discord bởi {ctx.author.name} ({ctx.author.id})")
                    await ctx.send("🛑 Đang tắt ứng dụng...")
                    
                    if self.app_shutdown_callback:
                        try:
                            self.app_shutdown_callback()
                        except Exception as e:
                            logger.error(f"Lỗi khi tắt ứng dụng: {str(e)}")
                            await ctx.send(f"❌ Lỗi khi tắt ứng dụng: {str(e)}")
                    else:
                        await ctx.send("⚠️ Callback tắt ứng dụng chưa được thiết lập")
                
                @self.bot.command(name='restart', aliases=['khoi_dong_lai', 'reload', 're'])
                async def restart_command(ctx):
                    """Lệnh khởi động lại ứng dụng"""
                    # Kiểm tra quyền
                    if ADMIN_USER_IDS and str(ctx.author.id) not in ADMIN_USER_IDS:
                        await ctx.send("❌ Bạn không có quyền sử dụng lệnh này!")
                        return
                    
                    logger.warning(f"⚠️ Lệnh khởi động lại ứng dụng từ Discord bởi {ctx.author.name} ({ctx.author.id})")
                    await ctx.send("🔄 Đang khởi động lại ứng dụng...")
                    
                    if self.app_restart_callback:
                        try:
                            self.app_restart_callback()
                        except Exception as e:
                            logger.error(f"Lỗi khi khởi động lại ứng dụng: {str(e)}")
                            await ctx.send(f"❌ Lỗi khi khởi động lại ứng dụng: {str(e)}")
                    else:
                        await ctx.send("⚠️ Callback khởi động lại ứng dụng chưa được thiết lập")
                
                @self.bot.command(name='status', aliases=['trạng thái', 'info'])
                async def status_command(ctx):
                    """Lệnh kiểm tra trạng thái ứng dụng"""
                    status_embed = discord.Embed(
                        title="📊 Trạng Thái Ứng Dụng",
                        color=0x4CAF50 if self.is_connected else 0xF44336
                    )
                    status_embed.add_field(
                        name="🤖 Bot Status",
                        value="✅ Đang hoạt động" if self.is_connected else "❌ Không kết nối",
                        inline=False
                    )
                    status_embed.add_field(
                        name="📡 Channel",
                        value=f"{self.channel.name}" if self.channel else "❌ Chưa có",
                        inline=True
                    )
                    status_embed.add_field(
                        name="🔔 Thông Báo",
                        value="✅ Bật" if ENABLE_DISCORD_NOTIFICATIONS else "❌ Tắt",
                        inline=True
                    )
                    status_embed.set_footer(text="Hệ thống Quản lý Quân nhân")
                    await ctx.send(embed=status_embed)
                
                @self.bot.command(name='help_bot', aliases=['h', 'commands'])
                async def help_command(ctx):
                    """Lệnh hiển thị danh sách lệnh"""
                    help_embed = discord.Embed(
                        title="📋 Danh Sách Lệnh",
                        description="Các lệnh điều khiển ứng dụng từ Discord",
                        color=0x2196F3
                    )
                    help_embed.add_field(
                        name="`!shutdown` hoặc `!tắt`",
                        value="Tắt ứng dụng",
                        inline=False
                    )
                    help_embed.add_field(
                        name="`!restart` hoặc `!khoi_dong_lai`",
                        value="Khởi động lại ứng dụng",
                        inline=False
                    )
                    help_embed.add_field(
                        name="`!status` hoặc `!trạng thái`",
                        value="Kiểm tra trạng thái ứng dụng",
                        inline=False
                    )
                    help_embed.add_field(
                        name="`!help_bot` hoặc `!h`",
                        value="Hiển thị danh sách lệnh này",
                        inline=False
                    )
                    help_embed.set_footer(text="Hệ thống Quản lý Quân nhân")
                    await ctx.send(embed=help_embed)
            
            @self.bot.event
            async def on_ready():
                logger.info("=" * 50)
                logger.info(f'✅ Discord bot đã kết nối: {self.bot.user}')
                logger.info(f'Bot ID: {self.bot.user.id}')
                logger.info(f'Bot đang ở {len(self.bot.guilds)} server(s):')
                for guild in self.bot.guilds:
                    logger.info(f'  - {guild.name} (ID: {guild.id})')
                logger.info("=" * 50)
                
                # Tìm channel để gửi thông báo
                if DISCORD_CHANNEL_ID:
                    try:
                        channel_id_int = int(DISCORD_CHANNEL_ID)
                        self.channel = self.bot.get_channel(channel_id_int)
                        if self.channel:
                            logger.info(f'✅ Đã tìm thấy channel theo ID: {self.channel.name} (ID: {DISCORD_CHANNEL_ID})')
                            logger.info(f'   Server: {self.channel.guild.name}')
                        else:
                            logger.warning(f'⚠️ Không tìm thấy channel với ID: {DISCORD_CHANNEL_ID}')
                            logger.info('Đang tìm tất cả channels...')
                            for guild in self.bot.guilds:
                                for ch in guild.text_channels:
                                    logger.info(f'   - {ch.name} (ID: {ch.id})')
                    except ValueError:
                        logger.error(f'❌ Channel ID không hợp lệ: {DISCORD_CHANNEL_ID}')
                    except Exception as e:
                        logger.error(f'❌ Lỗi khi tìm channel theo ID: {str(e)}', exc_info=True)
                
                # Nếu chưa có channel, tìm channel đầu tiên có quyền gửi tin nhắn
                if not self.channel:
                    logger.info('Đang tìm channel mặc định...')
                    for guild in self.bot.guilds:
                        logger.info(f'Đang tìm trong server: {guild.name}')
                        for channel in guild.text_channels:
                            try:
                                perms = channel.permissions_for(guild.me)
                                if perms.send_messages:
                                    self.channel = channel
                                    logger.info(f'✅ Đã tìm thấy channel mặc định: {channel.name} trong server {guild.name}')
                                    break
                                else:
                                    logger.debug(f'   Channel {channel.name} không có quyền gửi tin nhắn')
                            except Exception as e:
                                logger.error(f'Lỗi khi kiểm tra quyền channel {channel.name}: {str(e)}')
                        if self.channel:
                            break
                
                if self.channel:
                    logger.info(f'✅ Channel sẵn sàng: {self.channel.name} (ID: {self.channel.id})')
                    # Gửi thông báo test
                    try:
                        test_embed = discord.Embed(
                            title="🤖 Bot Đã Kết Nối",
                            description="Hệ thống Quản lý Quân nhân đã sẵn sàng!",
                            color=0x4CAF50
                        )
                        test_embed.set_footer(text="Hệ thống Quản lý Quân nhân")
                        await self.channel.send(embed=test_embed)
                        logger.info("✅ Đã gửi thông báo test thành công lên Discord")
                    except discord.errors.Forbidden as e:
                        logger.error(f"❌ Không có quyền gửi tin nhắn vào channel {self.channel.name}: {str(e)}")
                    except Exception as e:
                        logger.error(f"❌ Lỗi khi gửi thông báo test: {str(e)}", exc_info=True)
                else:
                    logger.error('❌ Không tìm thấy channel để gửi thông báo')
                    # Liệt kê tất cả channels để debug
                    for guild in self.bot.guilds:
                        logger.info(f'Channels trong server {guild.name}:')
                        for channel in guild.text_channels:
                            try:
                                perms = channel.permissions_for(guild.me)
                                can_send = "✅" if perms.send_messages else "❌"
                                logger.info(f'  {can_send} {channel.name} (ID: {channel.id})')
                            except:
                                logger.info(f'  ❓ {channel.name} (ID: {channel.id})')
                
                self.is_connected = True
                logger.info("✅ Bot đã sẵn sàng nhận và gửi thông báo")
                
                # Gửi các thông báo đã chờ trong queue
                queue_count = 0
                while not self.message_queue.empty():
                    try:
                        msg_data = self.message_queue.get_nowait()
                        asyncio.create_task(self._send_message(msg_data['message'], msg_data.get('embed')))
                        queue_count += 1
                    except queue.Empty:
                        break
                    except Exception as e:
                        logger.error(f"Lỗi khi gửi thông báo từ queue: {str(e)}")
                
                if queue_count > 0:
                    logger.info(f"Đã gửi {queue_count} thông báo từ queue")
            
            @self.bot.event
            async def on_error(event, *args, **kwargs):
                logger.error(f"Lỗi Discord bot event {event}: {args}, {kwargs}", exc_info=True)
            
            logger.info("Đang kết nối bot với Discord...")
            await self.bot.start(DISCORD_BOT_TOKEN)
        except discord.errors.LoginFailure as e:
            logger.error(f"❌ Lỗi đăng nhập Discord: Token không hợp lệ hoặc đã hết hạn")
            logger.error(f"Chi tiết: {str(e)}")
            self.is_connected = False
        except Exception as e:
            logger.error(f"❌ Lỗi khi khởi động Discord bot: {str(e)}", exc_info=True)
            self.is_connected = False
    
    def start(self):
        """Khởi động bot trong thread riêng"""
        if not DISCORD_AVAILABLE:
            logger.warning("Thư viện discord.py chưa được cài đặt")
            return
            
        if not DISCORD_BOT_TOKEN or not ENABLE_DISCORD_NOTIFICATIONS:
            return
        
        try:
            import threading
            self.loop = asyncio.new_event_loop()
            
            def run_bot():
                asyncio.set_event_loop(self.loop)
                self.loop.run_until_complete(self._start_bot())
            
            self.thread = threading.Thread(target=run_bot, daemon=True)
            self.thread.start()
            logger.info("Đã khởi động Discord bot thread")
        except Exception as e:
            logger.error(f"Lỗi khi khởi động Discord bot thread: {str(e)}")
    
    def stop(self):
        """Dừng bot"""
        if self.bot and self.loop:
            try:
                self.loop.call_soon_threadsafe(self.loop.stop)
                self.is_connected = False
                logger.info("Đã dừng Discord bot")
            except Exception as e:
                logger.error(f"Lỗi khi dừng Discord bot: {str(e)}")
    
    async def _send_message(self, message: str, embed: Optional[discord.Embed] = None):
        """Gửi tin nhắn lên Discord"""
        if not self.is_connected:
            logger.warning("Discord bot chưa kết nối")
            return False
            
        if not self.channel:
            logger.warning("Không có channel để gửi thông báo")
            return False
        
        try:
            if embed:
                await self.channel.send(message, embed=embed)
                logger.debug(f"Đã gửi embed: {embed.title}")
            else:
                await self.channel.send(message)
                logger.debug(f"Đã gửi message: {message[:50]}")
            return True
        except discord.errors.Forbidden as e:
            logger.error(f"Không có quyền gửi tin nhắn vào channel {self.channel.name}: {str(e)}")
            return False
        except Exception as e:
            logger.error(f"Lỗi khi gửi tin nhắn lên Discord: {str(e)}", exc_info=True)
            return False
    
    def send_notification(self, title: str, description: str, color: int = 0x4CAF50):
        """Gửi thông báo lên Discord (non-blocking)"""
        if not ENABLE_DISCORD_NOTIFICATIONS:
            logger.debug("Thông báo Discord đã bị tắt")
            return
        
        if not DISCORD_AVAILABLE:
            logger.warning("Thư viện discord.py chưa được cài đặt")
            return
        
        try:
            embed = discord.Embed(
                title=title,
                description=description,
                color=color
            )
            embed.set_footer(text="Hệ thống Quản lý Quân nhân")
            
            # Nếu bot chưa kết nối, thêm vào queue
            if not self.is_connected or not self.channel:
                logger.info(f"Bot chưa sẵn sàng, thêm thông báo vào queue: {title}")
                self.message_queue.put({
                    'message': '',
                    'embed': embed
                })
                # Đợi một chút để bot kết nối (tối đa 5 giây)
                for i in range(50):  # 50 lần x 0.1 giây = 5 giây
                    if self.is_connected and self.channel:
                        break
                    time.sleep(0.1)
            
            # Gửi trong event loop
            if self.loop and not self.loop.is_closed() and self.is_connected:
                try:
                    future = asyncio.run_coroutine_threadsafe(
                        self._send_message("", embed=embed),
                        self.loop
                    )
                    # Đợi tối đa 2 giây
                    future.result(timeout=2)
                    logger.info(f"✅ Đã gửi thông báo: {title}")
                except Exception as e:
                    logger.error(f"Lỗi khi gửi thông báo: {str(e)}")
                    # Thêm vào queue để thử lại sau
                    self.message_queue.put({
                        'message': '',
                        'embed': embed
                    })
            elif not self.is_connected:
                logger.warning(f"Bot chưa kết nối, thêm vào queue: {title}")
                self.message_queue.put({
                    'message': '',
                    'embed': embed
                })
        except Exception as e:
            logger.error(f"Lỗi khi gửi thông báo Discord: {str(e)}", exc_info=True)
    
    def notify_personnel_added(self, personnel_name: str):
        """Thông báo khi thêm quân nhân mới"""
        if NOTIFY_ON_PERSONNEL_ADD:
            self.send_notification(
                "➕ Thêm Quân Nhân Mới",
                f"Đã thêm quân nhân: **{personnel_name}**",
                color=0x4CAF50
            )
    
    def notify_personnel_updated(self, personnel_name: str):
        """Thông báo khi cập nhật quân nhân"""
        if NOTIFY_ON_PERSONNEL_UPDATE:
            self.send_notification(
                "✏️ Cập Nhật Quân Nhân",
                f"Đã cập nhật thông tin quân nhân: **{personnel_name}**",
                color=0xFF9800
            )
    
    def notify_personnel_deleted(self, personnel_name: str):
        """Thông báo khi xóa quân nhân"""
        if NOTIFY_ON_PERSONNEL_DELETE:
            self.send_notification(
                "🗑️ Xóa Quân Nhân",
                f"Đã xóa quân nhân: **{personnel_name}**",
                color=0xF44336
            )
    
    def notify_export(self, file_type: str, file_name: str, count: int = 0):
        """Thông báo khi xuất file"""
        if NOTIFY_ON_EXPORT:
            description = f"Đã xuất file **{file_type}**: `{file_name}`"
            if count > 0:
                description += f"\nSố lượng: {count} quân nhân"
            
            self.send_notification(
                "📄 Xuất File",
                description,
                color=0x2196F3
            )
    
    def notify_app_started(self, username: str = ""):
        """Thông báo khi ứng dụng khởi động"""
        if ENABLE_DISCORD_NOTIFICATIONS:
            description = "Hệ thống Quản lý Quân nhân đã được khởi động"
            if username:
                description += f"\nNgười dùng: **{username}**"
            
            self.send_notification(
                "🚀 Ứng Dụng Đã Khởi Động",
                description,
                color=0x4CAF50
            )
    
    def set_shutdown_callback(self, callback):
        """Thiết lập callback để tắt ứng dụng"""
        self.app_shutdown_callback = callback
    
    def set_restart_callback(self, callback):
        """Thiết lập callback để khởi động lại ứng dụng"""
        self.app_restart_callback = callback
    
    def test_connection(self):
        """Test kết nối bot và gửi thông báo test"""
        logger.info("=== TEST DISCORD BOT ===")
        logger.info(f"Discord available: {DISCORD_AVAILABLE}")
        logger.info(f"Token configured: {bool(DISCORD_BOT_TOKEN)}")
        logger.info(f"Notifications enabled: {ENABLE_DISCORD_NOTIFICATIONS}")
        logger.info(f"Remote control enabled: {ENABLE_REMOTE_CONTROL}")
        logger.info(f"Bot connected: {self.is_connected}")
        logger.info(f"Channel: {self.channel.name if self.channel else 'None'}")
        logger.info(f"Channel ID: {self.channel.id if self.channel else 'None'}")
        
        if self.is_connected and self.channel:
            self.send_notification(
                "🧪 Test Kết Nối",
                "Bot đang hoạt động bình thường!\nĐây là thông báo test từ hệ thống.\n\n"
                "**Các lệnh có sẵn:**\n"
                "`!help_bot` hoặc `!h` - Xem danh sách lệnh\n"
                "`!status` - Kiểm tra trạng thái\n"
                "`!shutdown` - Tắt ứng dụng\n"
                "`!restart` - Khởi động lại ứng dụng",
                color=0x2196F3
            )
            logger.info("✅ Đã gửi thông báo test")
            return True
        else:
            logger.warning("❌ Bot chưa kết nối hoặc chưa có channel")
            return False


# Singleton instance
_discord_bot_instance = None

def get_discord_bot() -> DiscordBotService:
    """Lấy instance của Discord bot service"""
    global _discord_bot_instance
    if _discord_bot_instance is None:
        _discord_bot_instance = DiscordBotService()
    return _discord_bot_instance

