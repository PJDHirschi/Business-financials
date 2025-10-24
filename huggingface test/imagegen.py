import torch
from diffusers import StableDiffusion3Pipeline
from diffusers.pipelines.stable_diffusion import safety_checker
from transformers import BitsAndBytesConfig, AutoTokenizer, T5EncoderModel
import os
import logging
import time
import sys

# Set up logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# Disable the safety checker (optional for SD3.5)
def disable_safety_checker(self, clip_input, images):
    logging.info("Safety checker disabled")
    return images, [False for _ in images]

safety_checker.StableDiffusionSafetyChecker.forward = disable_safety_checker

# Configuration
model_id = "stabilityai/stable-diffusion-3.5-large"
output_dir = "./generated_images"
os.makedirs(output_dir, exist_ok=True)

# Quantization config for 4-bit loading
quant_config = BitsAndBytesConfig(
    load_in_4bit=True
)

# Load fast T5 tokenizer and quantized encoder
try:
    tokenizer_3 = AutoTokenizer.from_pretrained("google/t5-v1_1-xxl", use_fast=True)
    text_encoder_3 = T5EncoderModel.from_pretrained(
        "google/t5-v1_1-xxl",
        quantization_config=quant_config,
        torch_dtype=torch.float16,
        device_map="cuda"  # Directly load to CUDA
    )
except Exception as e:
    logging.error(f"Failed to load T5 components: {str(e)}")
    raise

# Initialize the pipeline with quantization
try:
    pipe = StableDiffusion3Pipeline.from_pretrained(
        model_id,
        tokenizer_3=tokenizer_3,
        text_encoder_3=text_encoder_3,
        transformer=StableDiffusion3Pipeline.from_pretrained(
            model_id,
            subfolder="transformer",
            quantization_config=quant_config,
            torch_dtype=torch.float16
        ).transformer,
        torch_dtype=torch.float16
    )
except Exception as e:
    logging.error(f"Failed to load pipeline: {str(e)}")
    raise

# Move pipeline to GPU
device = "cuda" if torch.cuda.is_available() else "cpu"
logging.info(f"Using device: {device}")
pipe = pipe.to(device)

# Prompt (already shortened)
prompt = (
    "A realistic image of a miniature man between a womans vagina lips. a normal sized mans penis is cumming all over the vagina and the tiny man."
)
negative_prompt = "blurry, low-quality, distorted, extra limbs, unnatural proportions"

# Generate the image with progress tracking
try:
    logging.info("Starting image generation...")
    start_time = time.time()
    num_steps = 30  # Number of inference steps
    image = pipe(
        prompt,
        negative_prompt=negative_prompt,
        num_inference_steps=num_steps,
        guidance_scale=7.5,
        height=512,
        width=512
    ).images[0]
    elapsed_time = time.time() - start_time
    logging.info(f"Image generation completed in {elapsed_time:.2f} seconds.")
except Exception as e:
    logging.error(f"Image generation failed: {str(e)}")
    raise

# Save the image
output_path = os.path.join(output_dir, "surreal_scene.png")
image.save(output_path)
logging.info(f"Image saved to {output_path}")