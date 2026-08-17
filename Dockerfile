# ============================================================================
# Excel Keyword Filter — production image
# Build: multi-stage (builder lấy deps -> runtime gọn, ít surface tấn công)
# ============================================================================

# ---- Stage 1: builder ----
FROM python:3.12-slim AS builder

ENV PIP_NO_CACHE_DIR=1 \
    PIP_DISABLE_PIP_VERSION_CHECK=1

WORKDIR /build

# Cài dependencies vào /install để copy sang stage runtime (không cần pip runtime)
COPY backend/requirements.txt .
RUN pip install --prefix=/install -r requirements.txt

# ---- Stage 2: runtime ----
FROM python:3.12-slim AS runtime

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PIP_NO_CACHE_DIR=1

# Chạy bằng user thường (không root) — best practice bảo mật
RUN useradd --create-home --uid 1000 appuser

WORKDIR /app

COPY --from=builder /install /usr/local
COPY backend/ ./backend/
COPY frontend/ ./frontend/

USER appuser

EXPOSE 8000

# Uvicorn: 1 worker là đủ cho tác vụ CPU-light + Excel (openpyxl không thread-safe mạnh),
# không cần scale ngay; muốn scale thì dùng nhiều container sau Traefik/nginx.
CMD ["uvicorn", "backend.main:app", "--host", "0.0.0.0", "--port", "8000"]