from flask import Flask, request, jsonify, render_template
from flask_cors import CORS
from groq import Groq
import os
from werkzeug.utils import secure_filename
import pypdf
from pptx import Presentation
import json
from datetime import datetime

app = Flask(__name__, template_folder='templates')
CORS(app)

# ====================== API Key (مكتوب مباشرة) ======================
GROQ_API_KEY = "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"

client = Groq(api_key=GROQ_API_KEY)

# مجلد قاعدة المعرفة
LOCAL_KB_PATH = os.path.join(os.path.dirname(__file__), "knowledge_base")
os.makedirs(LOCAL_KB_PATH, exist_ok=True)


class SemanticNetwork:
    def __init__(self):
        self.concepts = {}
        self._initialize_core_knowledge()

    def _initialize_core_knowledge(self):
        core_concepts = [
            ("Integrated_Computer_Systems", "field", "Combinations of hardware, software, and networking components working cohesively to perform complex tasks."),
            ("Vertical_Integration", "method", "Combining systems within a specific domain or industry, like a single factory or hospital EHR."),
            ("Horizontal_Integration", "method", "Connecting systems across different domains or organizations, like a supply chain or smart city."),
            ("Data_Integration", "method", "Combining data from different sources ensuring seamless and consistent flow across applications."),
            ("Advanced_Systems_Software", "software", "Low-level programs like OS, device drivers, firmware (BIOS/UEFI), and utilities that manage hardware resources."),
            ("Von_Neumann_Architecture", "architecture", "Computer design (1945) with a single shared memory for both instructions and data, using sequential processing and one bus."),
            ("Harvard_Architecture", "architecture", "Computer design with separate memories and buses for instructions and data, allowing simultaneous, faster access. Used in DSPs."),
            ("Stored_Program_Concept", "principle", "Both data and instructions for a program are stored together in the computer's memory, eliminating the need to physically rewire machines."),
            ("Instruction_Set", "concept", "The vocabulary of the processor, defining all instructions it understands (OP-CODE, Source Operands, Result Operand)."),
            ("CISC", "architecture", "Complex Instruction Set Computer: complex hardware, simpler software, reduces instructions per program (e.g., Intel x86)."),
            ("RISC", "architecture", "Reduced Instruction Set Computer: simpler hardware, complex compiler, executes simple instructions very quickly in one cycle (e.g., ARM)."),
            ("Pipelining", "technique", "Overlapping the execution of instructions (Fetch, Decode, Execute) to reduce overall execution time and improve performance."),
            ("Data_Dependency", "pipeline_problem", "Occurs when adjacent instructions in a pipeline access the same memory location. Solved by Hardware Interlocks, Operand Forwarding, or Delayed Load."),
            ("Branch_Prediction", "pipeline_solution", "CPU guesses the path of program execution to avoid pipeline stalls (flushes) during branching."),
            ("Status_Flags", "hardware", "1-bit cells in the CPU PSW like Carry (CY), Parity (P), Zero (Z), Sign (S), and Overflow (V) flags."),
            ("Memory_Hierarchy", "architecture", "Arrangement of memory from fastest/smallest (Registers, Cache) to slowest/largest (Main Memory/RAM, Secondary Storage)."),
            ("SRAM", "hardware", "Static RAM: fast, expensive, used in cache memory, retains data without constant refreshing."),
            ("DRAM", "hardware", "Dynamic RAM: slower, cheaper, used in main memory, needs constant refreshing (every ~4ms)."),
            ("DDR2", "hardware", "Double Data Rate 2 RAM: transfers four pieces of data per clock cycle using multiplexing on a wider bus."),
            ("Endianness", "concept", "Order in which bytes are stored in memory: Big-Endian (most significant byte first) or Little-Endian (least significant byte first)."),
            ("Deadlock", "problem", "Occurs when threads wait indefinitely for resources held by each other (e.g., Dining Philosophers)."),
            ("Deadlock_Conditions", "rules", "Four conditions must happen simultaneously: Mutual Exclusion, Hold and Wait, No Preemption, and Circular Wait."),
            ("Deadlock_Prevention", "method", "Ensures at least one of the four deadlock conditions cannot hold by constraining how resource requests are made."),
            ("Deadlock_Avoidance", "method", "Dynamically avoids deadlocks using advance information about resource usage to ensure the system stays in a safe state (e.g., Banker's Algorithm)."),
            ("Process", "concept", "A program in execution, owning its own isolated memory space. Slower to create and requires IPC to communicate."),
            ("Thread", "concept", "The smallest unit of execution within a process. Faster to create and shares the parent process's memory space."),
            ("IPC", "method", "Inter-Process Communication: mechanisms like shared memory or message passing for isolated processes to coordinate safely."),
            ("Many_to_One_Model", "multithreading", "Many user threads controlled by a single kernel thread. Blocks completely if one thread waits for I/O."),
            ("One_to_One_Model", "multithreading", "Each user thread is directly mapped to a specific kernel thread, allowing true parallelism (used by Windows/Linux)."),
            ("Many_to_Many_Model", "multithreading", "User threads are mapped to a smaller or equal number of kernel threads dynamically for maximum efficiency without blocking."),
            ("System_SAAD", "concept", "A set of interrelated components that collect, process, store, and distribute information to support decision making and organizational control."),
            ("TPS", "IS_type", "Transaction Processing System: handles day-to-day routine business transactions."),
            ("MIS_Management", "IS_type", "Management Information System: provides structured reports to middle managers."),
            ("DSS", "IS_type", "Decision Support System: helps managers make semi-structured decisions."),
            ("ESS", "IS_type", "Executive Support System for senior executives."),
            ("ERP_System", "IS_type", "Enterprise Resource Planning system that integrates all departments."),
            ("SDLC", "methodology", "Systems Development Life Cycle: Planning, Analysis, Design, Implementation, Maintenance."),
            ("DFD", "modeling_tool", "Data Flow Diagram."),
            ("ERD", "modeling_tool", "Entity-Relationship Diagram."),
            ("Use_Case_Diagram", "modeling_tool", "Use Case Diagram."),
            ("Artificial_Intelligence", "field", "A branch of computer science devoted to creating machines that imitate human thinking."),
            ("Narrow_AI", "AI_type", "AI designed for a specific task only."),
            ("General_AI", "AI_type", "Hypothetical human-level AI."),
            ("Machine_Learning", "AI_technique", "Systems that learn from data without being explicitly programmed."),
            ("Deep_Learning", "AI_technique", "Machine Learning using multi-layered neural networks."),
            ("Natural_Language_Processing", "AI_application", "AI ability to understand and generate human language."),
            ("Computer_Vision", "AI_application", "AI field enabling machines to interpret visual data."),
            ("CIA_Triad", "security_principle", "Confidentiality, Integrity, Availability."),
            ("Firewall", "security_tool", "Network security device that monitors and filters traffic."),
            ("VPN", "security_tool", "Virtual Private Network for secure remote access."),
            ("Ransomware", "malware_type", "Malware that encrypts files and demands ransom."),
        ]
        for name, type_, desc in core_concepts:
            self.add_concept(name, type_, desc)

    def add_concept(self, name, concept_type, description):
        self.concepts[name] = {"type": concept_type, "description": description}

    def search(self, keywords):
        results = []
        for kw in keywords:
            kw_lower = kw.lower()
            for name, info in self.concepts.items():
                if kw_lower in name.lower() or kw_lower in info.get("description", "").lower():
                    if not any(r['name'] == name for r in results):
                        results.append({"name": name, **info})
        return results


class BrainBotExpertSystem:
    def __init__(self):
        self.kb = SemanticNetwork()
        self.history_folder = "chat_history"
        os.makedirs(self.history_folder, exist_ok=True)
        self.conversation_memory = {}
        self.load_local_files()

    def get_history_file_path(self, user_id):
        safe_id = "".join(c if c.isalnum() else "_" for c in user_id)
        return os.path.join(self.history_folder, f"{safe_id}_history.json")

    def load_history(self, user_id):
        filepath = self.get_history_file_path(user_id)
        if os.path.exists(filepath):
            try:
                with open(filepath, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except:
                return []
        return []

    def save_history(self, user_id, history):
        filepath = self.get_history_file_path(user_id)
        try:
            with open(filepath, 'w', encoding='utf-8') as f:
                json.dump(history, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"Error saving history for {user_id}: {e}")

    def get_user_memory(self, user_id):
        if user_id not in self.conversation_memory:
            self.conversation_memory[user_id] = self.load_history(user_id)
        return self.conversation_memory[user_id]

    def add_to_history(self, user_id, role, content):
        history = self.get_user_memory(user_id)
        history.append({
            "role": role,
            "content": content,
            "timestamp": datetime.now().isoformat()
        })
        self.save_history(user_id, history)

    def extract_concepts_with_ai(self, user_input):
        prompt = f"Extract the core computer science concepts from this text (fix typos automatically): '{user_input}'. Return ONLY a comma-separated list of 1 to 3 exact keywords. No chatting."
        try:
            completion = client.chat.completions.create(
                model="llama-3.1-8b-instant",
                messages=[{"role": "user", "content": prompt}],
                max_tokens=60,
                temperature=0.1
            )
            keywords = completion.choices[0].message.content.strip().split(',')
            return [k.strip() for k in keywords if k.strip()]
        except Exception:
            return [user_input]

    def load_local_files(self):
        for filename in os.listdir(LOCAL_KB_PATH):
            path = os.path.join(LOCAL_KB_PATH, filename)
            if os.path.isfile(path):
                self.process_file(path, filename)

    def process_file(self, path, filename):
        try:
            if filename.endswith(".txt"):
                with open(path, 'r', encoding='utf-8') as f:
                    self.ingest_text(f.read(), filename)
            elif filename.endswith(".pdf"):
                reader = pypdf.PdfReader(path)
                text = "\n".join([p.extract_text() or "" for p in reader.pages])
                self.ingest_text(text, filename)
            elif filename.endswith(".pptx"):
                prs = Presentation(path)
                text = ""
                for slide in prs.slides:
                    for shape in slide.shapes:
                        if hasattr(shape, "text"):
                            text += shape.text + " "
                self.ingest_text(text, filename)
        except Exception as e:
            print(f"Error processing {filename}: {e}")

    def ingest_text(self, text, source_name):
        chunks = [text[i:i + 3000] for i in range(0, len(text), 3000)]
        for i, chunk in enumerate(chunks):
            self.kb.add_concept(f"{source_name}_chunk{i}", "document", chunk[:2800])


brainbot = BrainBotExpertSystem()


@app.route('/')
def home():
    return render_template('index.html')


@app.route('/api/chat', methods=['POST'])
def chat():
    data = request.json
    user_message = data.get('message', '').strip()
    user_id = data.get('userId', 'student_1')

    if not user_message:
        return jsonify({'reply': "Bro, what do you want to ask? 😅"}), 400

    history = brainbot.get_user_memory(user_id)
    extracted_keywords = brainbot.extract_concepts_with_ai(user_message)
    matches = brainbot.kb.search(extracted_keywords)

    kb_context = "Relevant Knowledge:\n" + "\n".join([
        f"• {m['name']}: {m['description'][:280]}" for m in matches[:5]
    ]) if matches else "No specific KB match found."

    history_text = "\n".join([
        f"{m['role'].upper()}: {m['content']}" for m in history[-7:]
    ])

    system_prompt = "You are StudBud, a friendly, slightly sarcastic APU senior who helps students ace their exams. Speak naturally with Malaysian uni vibe. Keep most answers short and fun."

    user_prompt = f"""Context from Knowledge Base:
{kb_context}

Recent Conversation:
{history_text if history_text else "This is the start of conversation."}

Student: {user_message}

StudBud:"""

    try:
        completion = client.chat.completions.create(
            model="llama-3.1-8b-instant",
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_prompt}
            ],
            temperature=0.75,
            max_tokens=900
        )

        bot_reply = completion.choices[0].message.content.strip()

        brainbot.add_to_history(user_id, "user", user_message)
        brainbot.add_to_history(user_id, "assistant", bot_reply)

        return jsonify({'reply': bot_reply})

    except Exception as e:
        error_str = str(e).lower()
        if "rate limit" in error_str or "413" in error_str or "tokens" in error_str:
            return jsonify({'reply': "Bro, I hit the token limit 😂 Give me 30 seconds then try again."}), 429
        else:
            return jsonify({'reply': "Sorry bro, something went wrong. Try again."}), 500


@app.route('/api/history', methods=['GET'])
def get_history():
    user_id = request.args.get('userId', 'student_1')
    history = brainbot.get_user_memory(user_id)
    formatted_history = [{"role": msg["role"], "content": msg["content"]} for msg in history]
    return jsonify({"history": formatted_history})


@app.route('/api/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return jsonify({'error': 'No file'}), 400
    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': 'No selected file'}), 400

    filename = secure_filename(file.filename)
    save_path = os.path.join(LOCAL_KB_PATH, filename)
    file.save(save_path)
    brainbot.process_file(save_path, filename)
    return jsonify({'message': f'✅ {filename} added to StudBud knowledge base!'})


if __name__ == '__main__':
    print("StudBud Expert System is LIVE")
    print(f"Loaded {len(brainbot.kb.concepts)} concepts")
    app.run(host="0.0.0.0", port=int(os.getenv("PORT", 5000)), debug=False)