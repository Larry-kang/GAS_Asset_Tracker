/**
 * Util_Discord.js
 * Handles Discord Webhook integration with Hybrid Fallback logic.
 *
 * Protocols:
 * 1. Discord First: Try sending valid payload to Webhook.
 * 2. Email Fallback: If Discord fails or not configured, send to Admin Email.
 * 3. Silent: If neither is configured, log to system only.
 */

const Discord = {
    MAX_EMBED_DESCRIPTION_LENGTH: 3900,

    // Mapping level to colors (Decimal for Discord)
    COLORS: {
        INFO: 3447003,    // Blue
        WARNING: 16776960, // Yellow
        ERROR: 15158332,  // Red
        SUCCESS: 3066993, // Green
        STRATEGIC: 10181046 // Purple
    },

    /**
     * Send a systemic alert with hybrid fallback.
     * @param {string} title - Alert title
     * @param {string} description - Detailed message
     * @param {string} level - INFO, WARNING, ERROR, SUCCESS, STRATEGIC
     */
    sendAlert: function (title, description, level = 'INFO') {
        const webhookUrl = Settings.get('DISCORD_WEBHOOK_URL');
        const adminEmail = Settings.get('ADMIN_EMAIL');
        let discordSuccess = false;

        // 1. Try Discord
        if (webhookUrl) {
            discordSuccess = this._sendToDiscord(webhookUrl, title, description, level);
        } else {
            console.warn("Discord webhook is not configured.");
        }

        // 2. Fallback to Email if Discord failed or not set, AND level is critical enough
        // We send email for ERROR/WARNING/STRATEGIC if Discord is missing/failed.
        if (!discordSuccess && adminEmail) {
            // Only fallback for important messages to avoid spam
            if (['ERROR', 'WARNING', 'STRATEGIC'].includes(level)) {
                this._sendToEmail(adminEmail, title, description, level);
            }
        }

        return discordSuccess;
    },

    /**
     * Internal: Send payload to Discord Webhook
     * @private
     */
    _sendToDiscord: function (url, title, description, level) {
        try {
            // Prevent Rate Limit Abuse (Simple generic safeguard)
            Utilities.sleep(100);

            const color = this.COLORS[level] || this.COLORS.INFO;
            const safeDescription = this._truncateEmbedDescription(description);

            const payload = {
                embeds: [{
                    title: title,
                    description: safeDescription,
                    color: color,
                    timestamp: new Date().toISOString(),
                    footer: {
                        text: "SAP Command Center"
                    }
                }]
            };

            const options = {
                method: 'post',
                contentType: 'application/json',
                payload: JSON.stringify(payload),
                muteHttpExceptions: true
            };

            const response = UrlFetchApp.fetch(url, options);
            // Discord returns 204 No Content on success
            if (response.getResponseCode() === 204) {
                console.log("Discord webhook delivered successfully.");
                return true;
            } else {
                console.error(`Discord JSON Error (${response.getResponseCode()}): ${response.getContentText()}`);
                return false;
            }
        } catch (e) {
            console.error(`Discord Webhook Failed: ${e.toString()}`);
            return false;
        }
    },

    _truncateEmbedDescription: function (description) {
        const text = String(description || "");
        const maxLength = this.MAX_EMBED_DESCRIPTION_LENGTH;
        if (text.length <= maxLength) return text;

        const suffix = "\n\n...[Discord 訊息過長，已截斷；完整內容請看 Email 或 GAS 報告]";
        return text.substring(0, maxLength - suffix.length) + suffix;
    },

    /**
     * Internal: Fallback Email Sender
     * @private
     */
    _sendToEmail: function (email, title, body, level) {
        try {
            MailApp.sendEmail({
                to: email,
                subject: `[SAP ${level}] ${title}`,
                body: `${body}\n\n(Sent via SAP Fallback Notification)`
            });
            console.log(`Fallback email sent to ${email}`);
        } catch (e) {
            console.error(`Email Fallback Failed: ${e.toString()}`);
        }
    }
};

function sendDiscordAlert_(title, description, level) {
    return Discord.sendAlert(title, description, level || 'INFO');
}

function testDiscordNotification() {
    const ui = SpreadsheetApp.getUi();
    const webhookUrl = Settings.get('DISCORD_WEBHOOK_URL');
    if (!webhookUrl) {
        ui.alert('Discord 測試失敗', 'DISCORD_WEBHOOK_URL 尚未設定。請先執行「設定 Discord Webhook」。', ui.ButtonSet.OK);
        return;
    }

    const ok = sendDiscordAlert_(
        'SAP Discord 測試',
        '這是一則測試訊息。若你看到這則訊息，代表 GAS 到 Discord Webhook 的通道正常。',
        'INFO'
    );

    ui.alert(
        ok ? 'Discord 測試成功' : 'Discord 測試失敗',
        ok
            ? 'Discord webhook 已成功回應。'
            : 'Discord webhook 未成功回應。請查看 Apps Script 執行記錄中的 Discord JSON Error / Webhook Failed。',
        ui.ButtonSet.OK
    );
}
