import torch
print("CUDA available:", torch.cuda.is_available())
if torch.cuda.is_available():
    print("GPU:", torch.cuda.get_device_name(0))
    print("torch cuda:", torch.version.cuda)
    print("cudnn:", torch.backends.cudnn.version())