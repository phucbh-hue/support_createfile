import Anthropic from '@anthropic-ai/sdk';
import { authorize } from '../lib/auth.js';

const MODEL = process.env.CLAUDE_MODEL || 'claude-opus-5-5';
const EFFORT = process.env.CLAUDE_EFFORT || 'medium';
const MAX_INPUT_CHARS = 600_000;

const SYSTEM = `Bạn hỗ trợ nhân viên vận hành điền mẫu Word hàng loạt từ file dữ liệu. Mỗi yêu cầu gồm: (1) bản mô tả cấu trúc một mẫu .docx, trong đó mỗi đoạn văn, dòng bảng và ô đều có ID; (2) nội dung file dữ liệu đã chuyển thành văn bản: với Excel/CSV, mỗi ô ghi dạng CỘT=giá trị theo từng dòng R<số dòng> trong từng sheet; với file Word, mỗi bảng ghi theo cùng kiểu R<số dòng> và CỘT=giá trị (ô xuống dòng nối bằng " / "), còn đoạn văn ngoài bảng ghi dạng Đ<số>: nội dung (dấu " | " là phím Tab) - danh sách có thể nằm trong bảng hoặc trong các đoạn văn; (3) mô tả của người dùng về việc cần làm, kèm vài gợi ý tùy chọn.

Bạn trả về một kế hoạch JSON. Một chương trình sẽ áp kế hoạch đó lên mẫu một cách máy móc để sinh ra từng file, nên bạn không viết XML và không tự soạn lại văn bản của mẫu. Người dùng sẽ ký và gửi các file này cho khách hàng, vì vậy độ chính xác của tên, số điện thoại, địa chỉ, số lượng và số tiền quan trọng hơn mọi thứ khác.

# Cách chương trình áp kế hoạch

Với mỗi phần tử trong "documents", chương trình lấy một bản sao mới của mẫu rồi làm lần lượt:

1. Chạy "setup" (giống nhau cho mọi file). Dùng khi mẫu thiếu chỗ để điền:
   - clone_row: nhân bản dòng bảng "target", chèn ngay sau nó, dòng mới có ID là "as". Ô và đoạn văn trong dòng mới có ID "<as>.c0", "<as>.c0.p0"...
   - insert_paragraph: chèn một đoạn văn trống ngay sau đoạn "target", sao chép định dạng từ đoạn "copy_from" (để trống thì lấy định dạng của "target"), đoạn mới có ID là "as". Ví dụ: khối chữ ký của một bên chưa có dòng ghi tên trong khi bên kia có, thì chèn thêm đoạn để tên hai bên ngang hàng nhau.
   Mảng rỗng nếu không cần.
2. Điền "slots". Mỗi slot gắn một "key" với một đoạn văn "target" (ID có dạng p3 hoặc t0.r1.c3.p0). "template" là toàn bộ nội dung mới của đoạn văn đó, với {value} là chỗ đặt giá trị riêng của từng file. Hãy bắt đầu template bằng đúng phần chữ sẵn có của đoạn (nhãn, khoảng trắng đầu dòng) như trong bản mô tả, ví dụ "Người nhận: {value}" hoặc "Thành tiền bằng chữ: {value}./.", vì phần trùng với nội dung cũ sẽ giữ nguyên định dạng. Ô trống thì template là "{value}". "align" để "keep", trừ khi ô đang trống và nên canh lại (ví dụ số liệu trong bảng canh giữa giống các ô bên cạnh). "bold" chỉ bật khi giá trị cần in đậm mà đoạn đích chưa đậm.
3. Xóa các ID trong "remove" của file đó (dòng bảng, đoạn văn hoặc cả bảng). Dùng khi một dòng của mẫu không áp dụng cho file này, ví dụ dòng sản phẩm có số lượng bằng 0. Nếu xóa dòng làm lệch số thứ tự thì điền lại cột STT qua slot.

Trong một file, slot nào không có giá trị trong "values" thì đoạn văn đó giữ nguyên như mẫu. Vì vậy chỉ tạo slot cho những chỗ thật sự cần thay đổi và đừng đụng tới nội dung cố định của mẫu.

# Số tiền và phép tính

Đừng tự tính nhẩm. Trong "value" hãy viết biểu thức để chương trình tính, chỉ dùng số và + - * / ( ):
- {{num: 35*200000}} cho ra 7.000.000 (phân cách hàng nghìn bằng dấu chấm)
- {{Words: 35*200000 + 50*500000}} cho ra "Ba mươi hai triệu" (đọc số bằng chữ tiếng Việt, viết hoa chữ đầu; {{words: ...}} thì viết thường). Tự thêm đơn vị phía sau nếu cần, ví dụ "{{Words: 7000000}} đồng".
Số trong biểu thức viết liền, không có dấu phân cách. Các giá trị khác (tên, địa chỉ, số điện thoại) chép đúng nguyên văn từ file dữ liệu, chỉ gọn lại khoảng trắng thừa; giữ số 0 đầu của số điện thoại.

# Cách quyết định

- Làm theo mô tả của người dùng. Chỗ nào mô tả không nói rõ thì suy ra từ cấu trúc mẫu và dữ liệu, chọn cách hợp lý nhất và ghi giả định đó vào "notes".
- Chỉ dùng dữ liệu có trong file. Thiếu dữ liệu cho một chỗ cần điền thì để nguyên như mẫu và nêu trong "notes", không tự bịa.
- Bỏ qua các dòng tổng cộng, tiêu đề, ghi chú trong file dữ liệu; chúng không phải là một bản ghi.
- "Số file dự kiến" và "số khách hàng" người dùng nhập chỉ là ước lượng và có thể sai. Hãy bám theo dữ liệu thật; nếu kết quả lệch với con số họ nhập thì giải thích ngắn gọn lý do trong "notes" (ví dụ một người không có mục nào nên không sinh file).
- "Các trường cần lấy" cho biết người dùng muốn những cột nào xuất hiện trong file; điền chúng vào vị trí phù hợp nhất của mẫu.
- Không sinh file rỗng: nếu một bản ghi không có gì để điền theo yêu cầu thì bỏ qua và nêu trong "notes".
- "filename" ngắn gọn, phân biệt được từng file, có dấu tiếng Việt bình thường, không cần đuôi .docx.

"notes" viết bằng tiếng Việt, ngắn gọn, dành cho người dùng đọc trước khi tải file: các giả định đã chọn, bản ghi bị bỏ qua và lý do, chỗ lệch so với con số dự kiến, những điểm họ nên kiểm tra lại. Dùng dấu "-" thay cho gạch ngang dài.`;

const str = { type: 'string' };
const PLAN_SCHEMA = {
  type: 'object',
  additionalProperties: false,
  required: ['notes', 'setup', 'slots', 'documents'],
  properties: {
    notes: str,
    setup: {
      type: 'array',
      items: {
        type: 'object',
        additionalProperties: false,
        required: ['op', 'target', 'copy_from', 'as'],
        properties: { op: { type: 'string', enum: ['clone_row', 'insert_paragraph'] }, target: str, copy_from: str, as: str },
      },
    },
    slots: {
      type: 'array',
      items: {
        type: 'object',
        additionalProperties: false,
        required: ['key', 'target', 'template', 'align', 'bold'],
        properties: { key: str, target: str, template: str, align: { type: 'string', enum: ['keep', 'left', 'center', 'right'] }, bold: { type: 'boolean' } },
      },
    },
    documents: {
      type: 'array',
      items: {
        type: 'object',
        additionalProperties: false,
        required: ['filename', 'values', 'remove'],
        properties: {
          filename: str,
          values: {
            type: 'array',
            items: { type: 'object', additionalProperties: false, required: ['key', 'value'], properties: { key: str, value: str } },
          },
          remove: { type: 'array', items: str },
        },
      },
    },
  },
};

function buildUserMessage(body) {
  const hints = [];
  if (body.expectedFiles) hints.push(`- Số file dự kiến (ước lượng, có thể sai): ${body.expectedFiles}`);
  if (body.customerCount) hints.push(`- Số khách hàng / bản ghi (ước lượng, có thể sai): ${body.customerCount}`);
  if (body.fields) hints.push(`- Các trường cần lấy: ${body.fields}`);
  return [
    `<mau_word ten="${body.templateName || 'template.docx'}">\n${body.template}\n</mau_word>`,
    `<du_lieu ten="${body.dataName || 'data'}">\n${body.data}\n</du_lieu>`,
    `<mo_ta_cua_nguoi_dung>\n${body.prompt}\n</mo_ta_cua_nguoi_dung>`,
    hints.length ? `<goi_y>\n${hints.join('\n')}\n</goi_y>` : '',
    `Hôm nay là ${new Date().toLocaleDateString('vi-VN', { timeZone: 'Asia/Ho_Chi_Minh', day: '2-digit', month: '2-digit', year: 'numeric' })}. Hãy lập kế hoạch điền mẫu.`,
  ].filter(Boolean).join('\n\n');
}

function describeError(error) {
  if (error instanceof Anthropic.AuthenticationError) return 'API key Claude không hợp lệ hoặc đã bị thu hồi. Kiểm tra biến ANTHROPIC_API_KEY trên Vercel.';
  if (error instanceof Anthropic.PermissionDeniedError) return 'API key không có quyền dùng model này.';
  if (error instanceof Anthropic.RateLimitError) return 'Claude API đang giới hạn tần suất. Chờ một lát rồi thử lại.';
  if (error instanceof Anthropic.BadRequestError) return `Claude API từ chối yêu cầu: ${error.message}`;
  if (error instanceof Anthropic.APIConnectionError) return 'Không kết nối được tới Claude API.';
  if (error instanceof Anthropic.APIError) return `Lỗi Claude API (${error.status}): ${error.message}`;
  return `Lỗi máy chủ: ${error?.message || error}`;
}

export default async function handler(req, res) {
  if (req.method !== 'POST') return res.status(405).json({ error: 'Chỉ hỗ trợ POST.' });
  const access = authorize(req);
  if (!access.ok) return res.status(401).json({ error: access.error });
  if (!process.env.ANTHROPIC_API_KEY) return res.status(500).json({ error: 'Máy chủ chưa cấu hình ANTHROPIC_API_KEY.' });

  const body = req.body || {};
  for (const field of ['template', 'data', 'prompt']) {
    if (typeof body[field] !== 'string' || !body[field].trim()) return res.status(400).json({ error: `Thiếu nội dung "${field}".` });
  }
  if (body.template.length + body.data.length + body.prompt.length > MAX_INPUT_CHARS) {
    return res.status(413).json({ error: 'Mẫu và dữ liệu quá lớn cho một lần xử lý. Hãy tách file dữ liệu thành nhiều phần nhỏ hơn.' });
  }

  res.writeHead(200, { 'Content-Type': 'application/x-ndjson; charset=utf-8', 'Cache-Control': 'no-cache, no-transform', 'X-Accel-Buffering': 'no' });
  const send = (event) => res.write(JSON.stringify(event) + '\n');
  const ping = setInterval(() => send({ type: 'ping' }), 10_000);
  const started = Date.now();
  // Log vận hành: không ghi nội dung mẫu hay dữ liệu khách hàng.
  const log = (outcome, extra = {}) => console.log(JSON.stringify({ event: 'generate', outcome, seconds: Math.round((Date.now() - started) / 1000), inputChars: body.template.length + body.data.length, ...extra }));

  try {
    const client = new Anthropic();
    // Tự chuyển sang model dự phòng nếu bộ lọc an toàn từ chối nhầm yêu cầu.
    const stream = client.beta.messages.stream({
      model: MODEL,
      max_tokens: 64000,
      betas: ['server-side-fallback-2026-07-01'],
      fallbacks: 'default',
      system: SYSTEM,
      output_config: { effort: EFFORT, format: { type: 'json_schema', schema: PLAN_SCHEMA } },
      messages: [{ role: 'user', content: buildUserMessage(body) }],
    });
    req.on('close', () => { if (!res.writableEnded) stream.abort(); });
    send({ type: 'start', model: MODEL });
    for await (const event of stream) {
      if (event.type === 'content_block_start' && event.content_block.type === 'thinking') send({ type: 'thinking' });
      else if (event.type === 'content_block_delta' && event.delta.type === 'text_delta') send({ type: 'delta', text: event.delta.text });
    }
    const message = await stream.finalMessage();
    log(message.stop_reason, { model: message.model, inputTokens: message.usage.input_tokens, outputTokens: message.usage.output_tokens });
    if (message.stop_reason === 'refusal') send({ type: 'error', message: 'Claude từ chối xử lý yêu cầu này.' });
    else if (message.stop_reason === 'max_tokens') send({ type: 'error', message: 'Kết quả quá dài và bị cắt giữa chừng. Hãy tách file dữ liệu thành nhiều phần nhỏ hơn.' });
    else send({ type: 'done', model: message.model, usage: { input: message.usage.input_tokens, output: message.usage.output_tokens } });
  } catch (error) {
    console.error(error);
    log('error', { status: error?.status });
    send({ type: 'error', message: describeError(error) });
  } finally {
    clearInterval(ping);
    res.end();
  }
}
