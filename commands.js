// 社内ドメイン（これ以外は外部扱い）
const INTERNAL_DOMAIN = "jasty-bim.com";
const MAX_LEN = 500; // Smart Alertsのメッセージ上限

Office.onReady();

function call(fn) {
  return new Promise((resolve, reject) => {
    fn((r) => {
      if (r.status === Office.AsyncResultStatus.Succeeded) resolve(r.value);
      else reject(r.error);
    });
  });
}

function isExternal(address) {
  const a = (address || "").trim().toLowerCase();
  const at = a.lastIndexOf("@");
  if (at < 0) return false; // 解決できないアドレスは判定対象外
  return a.slice(at + 1) !== INTERNAL_DOMAIN;
}

function joinLimited(list, limit) {
  let out = "";
  for (let i = 0; i < list.length; i++) {
    const next = (out ? ", " : "") + list[i];
    if ((out + next).length > limit) return out + ` …他${list.length - i}件`;
    out += next;
  }
  return out;
}

async function onMessageSendHandler(event) {
  try {
    const item = Office.context.mailbox.item;
    const [to, cc, bcc, atts] = await Promise.all([
      call((cb) => item.to.getAsync(cb)),
      call((cb) => item.cc.getAsync(cb)),
      call((cb) => item.bcc.getAsync(cb)),
      call((cb) => item.getAttachmentsAsync(cb)),
    ]);

    const externals = [...to, ...cc, ...bcc]
      .map((r) => r.emailAddress)
      .filter(isExternal);
    // 署名などのインライン画像は除外
    const files = atts.filter((a) => !a.isInline).map((a) => a.name);

    if (externals.length === 0 || files.length === 0) {
      event.completed({ allowEvent: true });
      return;
    }

    const head = "外部宛先に添付ファイル付きで送信します。\n";
    const foot = "\n宛先と添付ファイルに誤りがなければ「送信」を選択してください。";
    const budget = MAX_LEN - head.length - foot.length - 30;
    const msg =
      head +
      `【外部宛先】${joinLimited(externals, Math.floor(budget / 2))}\n` +
      `【添付】${joinLimited(files, Math.floor(budget / 2))}` +
      foot;

    event.completed({
      allowEvent: false,
      errorMessage: msg.slice(0, MAX_LEN),
    });
  } catch (e) {
    // 取得に失敗した場合は送信を妨げない
    event.completed({ allowEvent: true });
  }
}

Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
