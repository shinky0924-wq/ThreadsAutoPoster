/**
 * Spintax and dynamic content processing utilities
 */

export function parseSpintax(text: string): string {
  if (!text) return '';
  let result = text;
  const spintaxRegex = /\{([^{}]+)\}/g;
  
  // Recursively resolve nested or consecutive spintax blocks
  let safetyLimit = 10;
  while (spintaxRegex.test(result) && safetyLimit > 0) {
    result = result.replace(spintaxRegex, (_match, group) => {
      const choices = group.split('|');
      const pick = choices[Math.floor(Math.random() * choices.length)];
      return pick.trim();
    });
    safetyLimit--;
  }

  // Dynamic date/time placeholders
  const now = new Date();
  const year = now.getFullYear();
  const month = now.getMonth() + 1;
  const date = now.getDate();
  const weekdays = ['日曜日', '月曜日', '火曜日', '水曜日', '木曜日', '金曜日', '土曜日'];
  const weekday = weekdays[now.getDay()];

  result = result
    .replace(/\{today\}/gi, `${year}年${month}月${date}日(${weekday.charAt(0)})`)
    .replace(/\{date\}/gi, `${month}月${date}日`)
    .replace(/\{month\}/gi, `${month}月`)
    .replace(/\{day\}/gi, `${date}日`)
    .replace(/\{weekday\}/gi, weekday);

  return result;
}

/**
 * Extracts numeric user ID from auth_token / sessionid
 * Supports formats: "47094044745:HP477WPO..." or "59443198637%3AoWW..."
 */
export function extractUserIdFromToken(token: string): string {
  if (!token) return '';
  const decoded = decodeURIComponent(token.trim());
  if (decoded.includes(':')) {
    const parts = decoded.split(':');
    if (/^\d+$/.test(parts[0])) return parts[0];
  }
  if (token.includes('%3A') || token.includes('%3a')) {
    const parts = token.split(/%3[Aa]/);
    if (/^\d+$/.test(parts[0])) return parts[0];
  }
  return '';
}

/**
 * Resolves image path from cell or detects embedded image filename in post text
 */
export function resolveImagePath(
  rawImgPath?: string,
  postText?: string,
  customImages?: Record<string, string>
): string {
  if (rawImgPath && rawImgPath.trim()) {
    const clean = rawImgPath.trim();
    if (customImages) {
      const base = clean.split(/[\\/]/).pop() || clean;
      if (customImages[clean]) return customImages[clean];
      if (customImages[base]) return customImages[base];
    }
    if (clean.includes('image2.png')) return '/images/image2.png';
    if (clean.includes('image.png')) return '/images/image.png';
    if (clean.startsWith('http://') || clean.startsWith('https://') || clean.startsWith('data:')) {
      return clean;
    }
    return clean;
  }

  if (postText) {
    if (postText.includes('image2.png')) return '/images/image2.png';
    if (postText.includes('image.png')) return '/images/image.png';
  }

  return '';
}
