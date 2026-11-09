<?php
// Emails every lead from the research page (mechkar.html) to Shlomi,
// so a lead is not lost when the visitor doesn't press "send" in WhatsApp.

header('Content-Type: application/json; charset=utf-8');
date_default_timezone_set('Asia/Jerusalem');

if ($_SERVER['REQUEST_METHOD'] !== 'POST') {
    http_response_code(405);
    echo json_encode(['ok' => false]);
    exit;
}

$TO = 'shlomi@uniquetech.co.il';
$FROM = 'noreply@uniquetech.co.il';

function field($key, $max) {
    $v = isset($_POST[$key]) ? (string)$_POST[$key] : '';
    $v = trim($v);
    if (function_exists('mb_substr')) {
        $v = mb_substr($v, 0, $max, 'UTF-8');
    } else {
        $v = substr($v, 0, $max);
    }
    return $v;
}

function one_line($v) {
    return trim(preg_replace('/[\r\n\t]+/', ' ', $v));
}

// Honeypot: real visitors never see or fill this field; bots usually do.
if (field('website', 200) !== '') {
    echo json_encode(['ok' => true]);
    exit;
}

$name    = one_line(field('name', 120));
$company = one_line(field('company', 120));
$role    = one_line(field('role', 120));
$size    = one_line(field('size', 40));
$phone   = one_line(field('phone', 40));
$email   = one_line(field('email', 160));
$message = field('message', 2000);

if ($name === '' || ($phone === '' && $email === '')) {
    http_response_code(400);
    echo json_encode(['ok' => false]);
    exit;
}

// Light throttle: at most 5 leads per IP per hour.
$ip = isset($_SERVER['REMOTE_ADDR']) ? $_SERVER['REMOTE_ADDR'] : 'unknown';
$stamp = sys_get_temp_dir() . '/lead_' . md5($ip);
$recent = [];
if (is_readable($stamp)) {
    foreach (explode("\n", (string)@file_get_contents($stamp)) as $t) {
        if ($t !== '' && (int)$t > time() - 3600) { $recent[] = (int)$t; }
    }
}
if (count($recent) >= 5) {
    http_response_code(429);
    echo json_encode(['ok' => false]);
    exit;
}
$recent[] = time();
@file_put_contents($stamp, implode("\n", $recent));

$subjectText = 'פנייה חדשה מדף המחקר: ' . $name . ($company !== '' ? ' · ' . $company : '');
$subject = '=?UTF-8?B?' . base64_encode($subjectText) . '?=';

$body = "פנייה חדשה לפתרון תפור למחלקת כספים (מדף המחקר)\n\n"
      . "שם: $name\n"
      . "חברה: $company\n"
      . "תפקיד: $role\n"
      . "גודל צוות: $size\n"
      . "טלפון: $phone\n"
      . "אימייל: $email\n"
      . "מה חשוב לשפר: $message\n\n"
      . "נשלח: " . date('d/m/Y H:i') . "\n"
      . "הערה: ייתכן שהפונה לא לחץ \"שליחה\" בוואטסאפ, אז כדאי לחזור אליו גם אם לא קיבלת הודעה.\n";

$headers = [
    'MIME-Version: 1.0',
    'Content-Type: text/plain; charset=UTF-8',
    'Content-Transfer-Encoding: 8bit',
    'From: =?UTF-8?B?' . base64_encode('דף המחקר · The UNIQUE Way') . '?= <' . $FROM . '>',
];
if (filter_var($email, FILTER_VALIDATE_EMAIL)) {
    $headers[] = 'Reply-To: ' . $email;
}

$sent = @mail($TO, $subject, $body, implode("\r\n", $headers));

echo json_encode(['ok' => (bool)$sent]);
