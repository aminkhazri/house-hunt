/** Little bits of delight: ripples, confetti, toasts. All honor reduced motion. */

const reducedMotion = () => window.matchMedia("(prefers-reduced-motion: reduce)").matches;

/** Material-style ripple on any .btn, started from where the pointer landed. */
export function enableRipples(root = document) {
  root.addEventListener("pointerdown", (e) => {
    const btn = e.target.closest(".btn, .fab");
    if (!btn || reducedMotion()) return;
    const rect = btn.getBoundingClientRect();
    const size = Math.max(rect.width, rect.height);
    const dot = document.createElement("span");
    dot.className = "ripple";
    Object.assign(dot.style, {
      width: `${size}px`,
      height: `${size}px`,
      left: `${e.clientX - rect.left - size / 2}px`,
      top: `${e.clientY - rect.top - size / 2}px`,
    });
    btn.append(dot);
    dot.addEventListener("animationend", () => dot.remove(), { once: true });
  });
}

/** Burst of confetti from a point on screen (used when a home is added). */
export function confetti(x, y, count = 36) {
  if (reducedMotion()) return;
  const colors = ["#ff8f6b", "#ffd166", "#2f7a5b", "#7cc6fe", "#c3a3ff"];
  for (let i = 0; i < count; i++) {
    const piece = document.createElement("span");
    piece.className = "confetti";
    piece.style.background = colors[i % colors.length];
    piece.style.left = `${x}px`;
    piece.style.top = `${y}px`;
    document.body.append(piece);

    const angle = Math.random() * Math.PI * 2;
    const speed = 120 + Math.random() * 220;
    const dx = Math.cos(angle) * speed;
    const dy = Math.sin(angle) * speed - 160;
    piece.animate(
      [
        { transform: "translate(0,0) rotate(0deg)", opacity: 1 },
        { transform: `translate(${dx}px, ${dy + 380}px) rotate(${Math.random() * 720}deg)`, opacity: 0 },
      ],
      { duration: 900 + Math.random() * 600, easing: "cubic-bezier(.2,.7,.4,1)" },
    ).finished.then(() => piece.remove());
  }
}

/**
 * Show a toast. Pass an action ({label, onClick}) for things like Undo.
 * Returns a function that dismisses it early.
 */
export function toast(message, { action, timeout = 5000 } = {}) {
  const host = document.getElementById("toasts");
  const el = document.createElement("div");
  el.className = "toast";
  const text = document.createElement("span");
  text.textContent = message;
  el.append(text);

  let timer;
  const close = () => {
    clearTimeout(timer);
    el.classList.add("is-leaving");
    el.addEventListener("animationend", () => el.remove(), { once: true });
    if (reducedMotion()) el.remove();
  };

  if (action) {
    const btn = document.createElement("button");
    btn.type = "button";
    btn.textContent = action.label;
    btn.addEventListener("click", () => { close(); action.onClick(); });
    el.append(btn);
  }
  host.append(el);
  timer = setTimeout(close, timeout);
  return close;
}
