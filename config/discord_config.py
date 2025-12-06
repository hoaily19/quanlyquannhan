"""
Cấu hình Discord Bot
Lưu ý: Token này nên được bảo vệ và không chia sẻ công khai
"""

# Discord Bot Token
DISCORD_BOT_TOKEN = "MTQ0NjgyMTQ2MTYwNzE4NjUwMw.Gr82Za.d9LHBBY54-PS3ys4hiLF5UYL3k-HILnjmeg3O8"

# Channel ID để gửi thông báo (có thể để None nếu muốn gửi vào channel mặc định)
# Lấy Channel ID: Bật Developer Mode trong Discord > Right-click channel > Copy ID
DISCORD_CHANNEL_ID = "1446826286004572210"  # Channel ID từ URL: discord.com/channels/.../1446826286004572210

# Cấu hình thông báo
ENABLE_DISCORD_NOTIFICATIONS = True  # Bật/tắt thông báo Discord
NOTIFY_ON_PERSONNEL_ADD = True  # Thông báo khi thêm quân nhân
NOTIFY_ON_PERSONNEL_UPDATE = True  # Thông báo khi cập nhật quân nhân
NOTIFY_ON_PERSONNEL_DELETE = True  # Thông báo khi xóa quân nhân
NOTIFY_ON_EXPORT = True  # Thông báo khi xuất file

# Cấu hình điều khiển từ xa
ENABLE_REMOTE_CONTROL = True  # Bật/tắt điều khiển từ xa qua Discord
ADMIN_USER_IDS = []  # Danh sách User ID được phép điều khiển (để trống = tất cả)

