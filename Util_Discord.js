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
        }

        // 2. Fallback to Email if Discord failed or not set, AND level is critical enough
        // We send email for ERROR/WARNING/STRATEGIC if Discord is missing/failed.
        if (!discordSuccess && adminEmail) {
            // Only fallback for important messages to avoid spam
            if (['ERROR', 'WARNING', 'STRATEGIC'].includes(level)) {
                this._sendToEmail(adminEmail, title, description, level);
            }
        }
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

            const payload = {
                embeds: [{
                    title: title,
                    description: description,
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
                return true;
            } else {
                console.error(`Discord JSON Error: ${response.getContentText()}`);
                return false;
            }
        } catch (e) {
            console.error(`Discord Webhook Failed: ${e.toString()}`);
            return false;
        }
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
