# ============================================
# SEP V4 Voice Engine — Text-to-Speech (TTS)
# Powered by Microsoft Edge TTS (edge-tts)
# ============================================

import asyncio
import io
import tempfile
import os
import base64
import hashlib
import time
from typing import Optional

import streamlit as st

# Voice options for Indonesian
VOICE_ID = "id-ID-ArdiNeural"       # Male voice — formal, clear
VOICE_ID_FEMALE = "id-ID-GadisNeural"  # Female voice — warm, natural
VOICE_ID_EN = "en-US-GuyNeural"     # English fallback

# Rate & pitch adjustments
RATE = "+10%"    # Slightly faster for presentations
PITCH = "+0Hz"   # Natural pitch


def _text_hash(text: str) -> str:
    """Short hash for caching audio files."""
    return hashlib.md5(text.encode()).hexdigest()[:12]


def synthesize_speech(
    text: str,
    voice: str = VOICE_ID,
    rate: str = RATE,
    pitch: str = PITCH,
) -> Optional[bytes]:
    """Synthesize speech and return MP3 bytes."""
    try:
        import edge_tts

        communicate = edge_tts.Communicate(
            text=text.strip(),
            voice=voice,
            rate=rate,
            pitch=pitch,
        )

        audio_data = asyncio.run(_collect_audio(communicate))
        return audio_data if audio_data else None

    except ImportError:
        st.warning("edge-tts belum terinstall. Run: pip install edge-tts")
        return None
    except Exception as e:
        st.warning(f"TTS error: {e}")
        return None


async def _collect_audio(communicate) -> bytes:
    """Async audio collection."""
    audio_chunks = []
    async for chunk in communicate.stream():
        if chunk["type"] == "audio":
            audio_chunks.append(chunk["data"])
    return b"".join(audio_chunks) if audio_chunks else b""


def get_cached_audio_path(text: str) -> str:
    """Get a predictable path for cached audio."""
    cache_dir = os.path.join(tempfile.gettempdir(), "sep_v4_tts_cache")
    os.makedirs(cache_dir, exist_ok=True)
    return os.path.join(cache_dir, f"sep_{_text_hash(text)}.mp3")


def get_or_generate_audio(text: str) -> Optional[str]:
    """Generate audio file, return path. Uses cache."""
    if not text or len(text.strip()) < 5:
        return None

    audio_path = get_cached_audio_path(text)

    # Return cached if recent (< 24 hours)
    if os.path.exists(audio_path):
        age = time.time() - os.path.getmtime(audio_path)
        if age < 86400:
            return audio_path

    # Generate new audio
    audio_bytes = synthesize_speech(text)
    if audio_bytes:
        with open(audio_path, "wb") as f:
            f.write(audio_bytes)
        return audio_path

    return None


def get_audio_base64(text: str) -> Optional[str]:
    """Generate audio and return as base64 data URI for HTML audio tag."""
    audio_bytes = synthesize_speech(text)
    if audio_bytes:
        b64 = base64.b64encode(audio_bytes).decode("utf-8")
        return f"data:audio/mp3;base64,{b64}"
    return None


def render_voice_button(narration: str, key_suffix: str = "") -> str:
    """Return HTML for a voice play button.

    Usage:
        st.markdown(render_voice_button(narration), unsafe_allow_html=True)
    """
    if not narration or len(narration.strip()) < 5:
        return ""

    audio_b64 = get_audio_base64(narration)
    if not audio_b64:
        return (
            '<span style="color:#94a3b8;font-size:13px;">'
            "🔇 Voice unavailable"
            "</span>"
        )

    unique_id = f"sep_voice_{_text_hash(narration + key_suffix)}"

    return f"""
    <div style="margin-top:10px;">
        <button
            onclick="
                var a = document.getElementById('{unique_id}');
                if(a.paused){{ a.play(); this.innerHTML='⏸️ Stop'; }}
                else{{ a.pause(); this.innerHTML='🔊 Dengarkan Narasi'; }}
            "
            style="
                background: linear-gradient(135deg, #6C5CE7, #A29BFE);
                color: white;
                border: none;
                border-radius: 50px;
                padding: 10px 22px;
                font-size: 14px;
                font-weight: 600;
                cursor: pointer;
                box-shadow: 0 4px 15px rgba(108,92,231,0.35);
                transition: all 0.2s;
            "
            onmouseover="this.style.transform='translateY(-2px)';this.style.boxShadow='0 6px 20px rgba(108,92,231,0.45)';"
            onmouseout="this.style.transform='translateY(0)';this.style.boxShadow='0 4px 15px rgba(108,92,231,0.35)';"
        >🔊 Dengarkan Narasi</button>
        <audio id="{unique_id}" src="{audio_b64}" style="display:none;"
            onended="
                var btn = document.querySelector('button[onclick*=\\\'{unique_id}\\\']');
                if(btn) btn.innerHTML='🔊 Dengarkan Narasi';
            ">
        </audio>
    </div>
    """


def render_voice_icon_button(narration: str, key_suffix: str = "") -> str:
    """Compact icon-only voice button."""
    if not narration or len(narration.strip()) < 5:
        return ""

    audio_b64 = get_audio_base64(narration)
    if not audio_b64:
        return ""

    unique_id = f"sep_voice_icon_{_text_hash(narration + key_suffix)}"

    return f"""
    <span style="display:inline-block;vertical-align:middle;margin-left:8px;">
        <button
            onclick="
                var a = document.getElementById('{unique_id}');
                if(a.paused){{ a.play(); this.textContent='⏸️'; }}
                else{{ a.pause(); this.textContent='🔊'; }}
            "
            style="
                background: linear-gradient(135deg, #6C5CE7, #A29BFE);
                color: white;
                border: none;
                border-radius: 50%;
                width: 36px;
                height: 36px;
                font-size: 15px;
                cursor: pointer;
                box-shadow: 0 2px 10px rgba(108,92,231,0.3);
            "
            onmouseover="this.style.transform='scale(1.1)';"
            onmouseout="this.style.transform='scale(1)';"
        >🔊</button>
        <audio id="{unique_id}" src="{audio_b64}" style="display:none;"></audio>
    </span>
    """
