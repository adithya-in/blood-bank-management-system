/* Dispatch Board — motion and lightweight interaction.
   Content is visible without JS; this file adds state feedback. */
(function () {
  "use strict";
  var reduce = window.matchMedia("(prefers-reduced-motion: reduce)").matches;
  var $ = function (sel, root) { return (root || document).querySelector(sel); };
  var $$ = function (sel, root) { return Array.prototype.slice.call((root || document).querySelectorAll(sel)); };

  /* ---------------------------------------------------------------- clock */
  var clock = $("[data-clock]");
  var dateEl = $("[data-date]");
  function tickClock() {
    var d = new Date();
    if (clock) {
      clock.textContent = String(d.getHours()).padStart(2, "0") + ":" + String(d.getMinutes()).padStart(2, "0");
    }
    if (dateEl) {
      dateEl.textContent = d.toLocaleDateString("en-GB", { weekday: "short", day: "2-digit", month: "short" }).toUpperCase();
    }
  }
  tickClock();
  setInterval(tickClock, 15000);

  /* --------------------------------------------------------------- toasts */
  function toast(message, kind) {
    var host = $(".toasts");
    if (!host) {
      host = document.createElement("div");
      host.className = "toasts";
      host.id = "toasts";
      document.body.appendChild(host);
    }
    var el = document.createElement("div");
    el.className = "toast toast-" + (kind || "success");
    el.setAttribute("role", "status");
    var mark = document.createElement("span");
    mark.textContent = kind === "error" ? "!" : "✓";
    mark.style.fontFamily = "var(--font-mono)";
    mark.style.fontWeight = "700";
    var text = document.createElement("span");
    text.textContent = message;
    var close = document.createElement("button");
    close.type = "button";
    close.setAttribute("aria-label", "Dismiss");
    close.textContent = "×";
    el.appendChild(mark); el.appendChild(text); el.appendChild(close);
    host.appendChild(el);
    function dismiss() {
      el.classList.add("is-leaving");
      setTimeout(function () { el.remove(); }, 240);
    }
    close.addEventListener("click", dismiss);
    setTimeout(dismiss, 4800);
  }

  $$(".toast").forEach(function (el) {
    var close = $("button", el);
    if (close) close.addEventListener("click", function () {
      el.classList.add("is-leaving");
      setTimeout(function () { el.remove(); }, 240);
    });
    setTimeout(function () {
      el.classList.add("is-leaving");
      setTimeout(function () { el.remove(); }, 240);
    }, 5200);
  });

  /* ----------------------------------------------------------- counters */
  function setCounter(el, value) {
    el.textContent = String(value);
    if (!reduce) {
      el.classList.remove("tick");
      void el.offsetWidth;
      el.classList.add("tick");
    }
  }

  function animateCounter(el, target) {
    if (reduce) { el.textContent = String(target); return; }
    var steps = 9;
    var i = 0;
    var timer = setInterval(function () {
      i += 1;
      if (i >= steps) {
        clearInterval(timer);
        setCounter(el, target);
      } else {
        var span = Math.max(2, Math.round(target * (1 - i / steps)));
        setCounter(el, Math.floor(Math.random() * span));
      }
    }, 46);
  }

  $$(".counter").forEach(function (el, idx) {
    var target = parseInt(el.getAttribute("data-value"), 10) || 0;
    if (reduce) { el.textContent = String(target); return; }
    el.textContent = "0";
    setTimeout(function () { animateCounter(el, target); }, 260 + idx * 55);
  });

  function animateCountNode(el) {
    var target = parseInt(el.getAttribute("data-count"), 10) || 0;
    if (reduce) { el.textContent = String(target); return; }
    var start = null;
    var from = 0;
    function step(ts) {
      if (!start) start = ts;
      var p = Math.min(1, (ts - start) / 720);
      var eased = 1 - Math.pow(1 - p, 3);
      el.textContent = String(Math.round(from + (target - from) * eased));
      if (p < 1) requestAnimationFrame(step);
      else el.textContent = String(target);
    }
    requestAnimationFrame(step);
  }
  $$("[data-count]").forEach(function (el) { animateCountNode(el); });

  /* stock bars */
  requestAnimationFrame(function () {
    $$("[data-bar]").forEach(function (bar, i) {
      var w = bar.style.getPropertyValue("--w") || "0%";
      setTimeout(function () { bar.style.width = w; }, 120 + i * 40);
    });
  });

  /* keep the active rail tab in view on narrow boards */
  (function () {
    var nav = $(".nav");
    var active = $(".nav a[aria-current='page']");
    if (!nav || !active || nav.scrollWidth <= nav.clientWidth) return;
    var left = active.offsetLeft - (nav.clientWidth - active.offsetWidth) / 2;
    nav.scrollLeft = Math.max(0, left);
  })();

  /* --------------------------------------------------------------- reveal */
  var reveals = $$(".reveal");
  if (reduce || !("IntersectionObserver" in window)) {
    reveals.forEach(function (el) { el.classList.add("is-in"); });
  } else {
    var io = new IntersectionObserver(function (entries) {
      entries.forEach(function (e) {
        if (e.isIntersecting) {
          e.target.classList.add("is-in");
          io.unobserve(e.target);
        }
      });
    }, { rootMargin: "0px 0px -8% 0px", threshold: 0.08 });
    reveals.forEach(function (el) { io.observe(el); });
  }

  /* ------------------------------------------------------ request status */
  function csrf() {
    var input = $('input[name="csrf_token"]');
    if (input) return input.value;
    var meta = $('meta[name="csrf-token"]');
    return meta ? meta.getAttribute("content") : "";
  }

  function updateStock(group, units) {
    var counter = $('.counter[data-group="' + group + '"]');
    if (counter) {
      counter.setAttribute("data-value", units);
      animateCounter(counter, units);
      var bar = counter.parentElement.querySelector("[data-bar]");
      if (bar) bar.style.width = Math.min(100, units) + "%";
    }
    $$("[data-count]").forEach(function (el) {
      if (el.closest(".stat") && el.closest(".stat").textContent.indexOf("Units on shelves") > -1) {
        animateCountNode(el);
      }
    });
  }

  function actionButtons(status) {
    if (status === "Pending") {
      return [["Approved", "Approve", ""], ["Cancelled", "Decline", " btn-danger-stamp"]];
    }
    if (status === "Approved") {
      return [["Completed", "Complete", ""], ["Cancelled", "Cancel", " btn-danger-stamp"]];
    }
    return null;
  }

  function wireRequestActions() {
    $$("[data-request]").forEach(function (card) {
      $$("[data-status]", card).forEach(function (btn) {
        if (btn.tagName !== "BUTTON") return;
        btn.addEventListener("click", function () {
          var id = card.getAttribute("data-request");
          var status = btn.getAttribute("data-status");
          var group = card.getAttribute("data-group") || "";
          btn.disabled = true;
          fetch("/api/requests/" + id + "/status", {
            method: "POST",
            headers: { "Content-Type": "application/json", "X-CSRF-Token": csrf() },
            body: JSON.stringify({ status: status })
          }).then(function (res) {
            return res.json().then(function (data) { return { ok: res.ok, data: data }; });
          }).then(function (result) {
            if (!result.ok) {
              btn.disabled = false;
              toast(result.data.error || "Could not update the request.", "error");
              return;
            }
            var slip = card.classList.contains("slip") ? card : null;
            if (slip) {
              var stamp = $(".slip-stamp", slip);
              if (stamp) {
                stamp.textContent = result.data.status === "Cancelled" ? "Declined" : result.data.status;
                stamp.style.color = result.data.status === "Cancelled" ? "var(--ink-soft)" : "";
              }
              slip.classList.add("is-stamping");
              setTimeout(function () {
                slip.style.transition = "opacity 260ms ease, transform 260ms ease";
                slip.style.opacity = "0";
                slip.style.transform = "translateY(-6px)";
                setTimeout(function () { slip.remove(); }, 300);
                var pending = $(".stat.is-alert .v [data-count]");
                if (pending) animateCountNode(pending);
              }, 1150);
              toast(result.data.status === "Cancelled" ? "Request declined." : "Request approved.", "success");
            } else {
              var chip = $("[data-status-chip]", card);
              if (chip) {
                chip.textContent = result.data.status;
                chip.className = "status status-" + result.data.status;
              }
              var holder = btn.closest(".row");
              if (holder) {
                holder.innerHTML = "";
                var next = actionButtons(result.data.status);
                if (next) {
                  next.forEach(function (spec) {
                    var b = document.createElement("button");
                    b.type = "button";
                    b.className = "btn btn-sm btn-stamp" + spec[2];
                    b.setAttribute("data-status", spec[0]);
                    b.textContent = spec[1];
                    holder.appendChild(b);
                  });
                } else {
                  var filed = document.createElement("span");
                  filed.className = "faint";
                  filed.style.fontSize = "11px";
                  filed.textContent = "Filed";
                  holder.appendChild(filed);
                }
              }
              card.style.transition = "background 900ms ease";
              card.style.background = "rgba(217,164,65,0.16)";
              setTimeout(function () { card.style.background = ""; }, 900);
              if (result.data.units !== null && result.data.units !== undefined && group) {
                updateStock(group, result.data.units);
              }
              toast("Request marked " + result.data.status.toLowerCase() + ".", "success");
              wireRequestActions();
            }
          }).catch(function () {
            btn.disabled = false;
            toast("Network error. Try again.", "error");
          });
        });
      });
    });
  }
  wireRequestActions();

  /* expose for debugging in the console */
  window.dispatchBoard = { toast: toast };
})();
