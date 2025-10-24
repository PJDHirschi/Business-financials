from transformers import AutoModelForCausalLM, AutoTokenizer
import torch
import sys

MODEL_NAME = "ystemsrx/Qwen2.5-Sex"   # or a local path that actually contains model files

TOP_P = 0.9
TOP_K = 80
TEMPERATURE = 0.3
MAX_NEW_TOKENS = 1080


# Pick device
if torch.cuda.is_available():
    device_map = "auto"
    torch_dtype = "auto"
else:
    device_map = {"": "cpu"}  # force CPU
    torch_dtype = "auto"

# Load
tokenizer = AutoTokenizer.from_pretrained(
    MODEL_NAME,
    trust_remote_code=True  # harmless if not needed; helpful if the model uses custom code
)

model = AutoModelForCausalLM.from_pretrained(
    MODEL_NAME,
    device_map=device_map,
    torch_dtype=torch_dtype,
    trust_remote_code=True
)

# Ensure pad/eos are defined
if tokenizer.pad_token_id is None and tokenizer.eos_token_id is not None:
    tokenizer.pad_token = tokenizer.eos_token

# Start messages (include a system prompt if you want)
messages = [{"role": "system", "content": ""}]

def chat_once(user_text: str) -> str:
    messages.append({"role": "user", "content": user_text})

    # Build chat-formatted text for Qwen
    text = tokenizer.apply_chat_template(
        messages,
        tokenize=False,
        add_generation_prompt=True
    )

    # Tokenize & to model device(s)
    inputs = tokenizer([text], return_tensors="pt")
    # If using device_map="auto", you can skip manual .to(device); the model will handle it.
    # But moving inputs to first device avoids CPU/GPU mismatch:
    inputs = {k: v.to(model.device) for k, v in inputs.items()}

    with torch.no_grad():
        generated = model.generate(
            **inputs,
            max_new_tokens=MAX_NEW_TOKENS,
            do_sample=True,
            top_p=TOP_P,
            top_k=TOP_K,
            temperature=TEMPERATURE,
            pad_token_id=tokenizer.pad_token_id,
            eos_token_id=tokenizer.eos_token_id
        )

    # Strip the prompt part to keep only the new tokens
    new_tokens = generated[0, inputs["input_ids"].shape[1]:]
    reply = tokenizer.decode(new_tokens, skip_special_tokens=True).strip()

    messages.append({"role": "assistant", "content": reply})
    return reply

if __name__ == "__main__":
    print("Loaded:", MODEL_NAME, "| Device:", model.device, "| torch:", torch.__version__)
    print("Type your message. Empty line to exit.\n")
    for line in sys.stdin:
        user_inp = line.strip()
        if not user_inp:
            break
        ans = chat_once(user_inp)
        print(f"Assistant: {ans}\n")
