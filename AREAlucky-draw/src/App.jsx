import React, { useCallback, useEffect, useMemo, useRef, useState } from "react";
import "./App.css";
import { ClipboardPaste, ChevronLeft, ChevronRight } from "lucide-react";
import * as XLSX from "xlsx";

const STORAGE_KEY = "lucky_draw_event_v1";

/* =========================
   AUDIO (ALWAYS ON AFTER USER GESTURE)
   ========================= */

function useAudioManager() {
  const SFX_VOL = 0.75;
  const SPIN_VOL = 0.35;
  const BGM_VOL = 0.28;

  const primedRef = useRef(false);
  const spinRef = useRef(null);
  const bgmRef = useRef(null);
  const oneShotPoolRef = useRef([]);

  // Load saved BGM volume from localStorage
  const [bgmVolume, setBgmVolume] = useState(() => {
    const saved = localStorage.getItem("lucky_draw_bgm_volume");
    return saved ? parseFloat(saved) : BGM_VOL;
  });

  // Save BGM volume to localStorage when it changes
  useEffect(() => {
    localStorage.setItem("lucky_draw_bgm_volume", bgmVolume.toString());
  }, [bgmVolume]);

  const prime = useCallback(() => {
    if (primedRef.current) return;
    primedRef.current = true;

    const spin = new Audio("/sfx/spin_loop.mp3");
    spin.loop = true;
    spin.volume = Math.max(0, Math.min(1, SPIN_VOL * SFX_VOL));
    spinRef.current = spin;

    const bgm = new Audio("/bgm/bgm.mp3");
    bgm.loop = true;
    bgm.volume = bgmVolume;
    bgmRef.current = bgm;
  }, [bgmVolume]);

  const ensureBgmPlaying = useCallback(() => {
    prime();
    const b = bgmRef.current;
    if (!b) return;
    b.volume = bgmVolume;
    b.play().catch(() => {});
  }, [prime, bgmVolume]);

  const updateBgmVolume = useCallback((newVolume) => {
    const clampedVolume = Math.max(0, Math.min(1, newVolume));
    setBgmVolume(clampedVolume);
    const b = bgmRef.current;
    if (b) {
      b.volume = clampedVolume;
    }
  }, []);

  const startSpin = useCallback(() => {
    prime();
    const s = spinRef.current;
    if (!s) return;
    s.volume = Math.max(0, Math.min(1, SPIN_VOL * SFX_VOL));
    if (!s.paused) return;
    s.currentTime = 0;
    s.play().catch(() => {});
  }, [prime]);

  const stopSpin = useCallback(() => {
    const s = spinRef.current;
    if (!s) return;
    s.pause();
    s.currentTime = 0;
  }, []);

  const getOneShot = useCallback(() => {
    // First try to find an available audio element
    let a = oneShotPoolRef.current.find((x) => x.paused || x.ended);
    if (!a) {
      // Limit pool size to prevent memory leaks during auto-draw
      if (oneShotPoolRef.current.length >= 5) {
        // Reuse the oldest ended audio element, or create space
        const endedIndex = oneShotPoolRef.current.findIndex((x) => x.ended);
        if (endedIndex >= 0) {
          a = oneShotPoolRef.current[endedIndex];
        } else {
          // If no ended elements, don't create more than 5
          return null;
        }
      } else {
        a = new Audio();
        oneShotPoolRef.current.push(a);
      }
    }
    return a;
  }, []);

  const playOneShot = useCallback(
    (src, vol = 1.0) => {
      prime();
      const a = getOneShot();
      if (!a) return; // Skip if no audio element available (pool full)
      a.src = src;
      a.currentTime = 0;
      a.volume = Math.max(0, Math.min(1, vol * SFX_VOL));
      a.play().catch(() => {});
    },
    [prime, getOneShot]
  );

  // Clean up ended audio elements to free memory
  const cleanupAudioPool = useCallback(() => {
    oneShotPoolRef.current = oneShotPoolRef.current.filter((a) => !a.ended);
  }, []);

  // ✅ IMPORTANT: memoize the returned object so `audio` doesn't change every render
  return useMemo(
    () => ({
      prime,
      ensureBgmPlaying,
      startSpin,
      stopSpin,
      bgmVolume,
      updateBgmVolume,
      cleanupAudioPool,
      click: () => playOneShot("/sfx/click.mp3", 0.9),
      confetti: () => playOneShot("/sfx/confetti.mp3", 1.0),

      // stop_tick is your "win/stop" sound now
      stopTick: (isAutoDraw = false) => playOneShot("/sfx/stop_tick.mp3", isAutoDraw ? 0.85 : 1.0),

      // kept for compatibility but NOT used by draw logic now
      win: (isAutoDraw = false) => playOneShot("/sfx/win.mp3", isAutoDraw ? 0.7 : 1.0),
    }),
    [prime, ensureBgmPlaying, startSpin, stopSpin, playOneShot, bgmVolume, updateBgmVolume, cleanupAudioPool]
  );
}

/* =========================
   UTIL
   ========================= */

const normalizeImageSrc = (raw) => {
  const s = (raw || "").trim();
  if (!s) return "";
  if (/^https?:\/\//i.test(s)) return s;
  if (/^data:image\//i.test(s)) return s;
  if (s.startsWith("/")) return s;
  return `/prizes/${s}`;
};

const buildPrizeWinnersMap = (winners) => {
  const map = {};
  (winners || []).forEach((w) => {
    const key = w.prize;
    map[key] = map[key] || [];
    map[key].push(w);
  });
  return map;
};

/* =========================
   DRAW PAGE (OUTSIDE)
   ========================= */

const DrawPage = ({
  audio,
  setCurrentPage,

  mode, // "union" | "general" | "money" | "bonus"
  prizes,
  participants,
  initialWinners = [],
  onWinnersChange,
  externalExcluded = [], // for money/bonus excludes round2 winners
  initialPrizeIndex = 0,
  onPrizeIndexChange,
  titleZh,
  titleEn,
  sheetName,
  moneyTotal,
  moneyWinnersCount,
}) => {
  const isMoneyLike = mode === "money" || mode === "bonus";

  const [currentPrizeIndex, setCurrentPrizeIndex] = useState(initialPrizeIndex);

  const [allRoundWinners, setAllRoundWinners] = useState(initialWinners);
  const [prizeWinners, setPrizeWinners] = useState(() => buildPrizeWinnersMap(initialWinners));

  const [absentWinners, setAbsentWinners] = useState([]);
  const [drawingState, setDrawingState] = useState("idle");
  const [currentWinner, setCurrentWinner] = useState(null);

  const [displayNumbers, setDisplayNumbers] = useState([0, 0, 0]);
  const [undoStack, setUndoStack] = useState([]);

  const [allSpinning, setAllSpinning] = useState(false);
  const [autoDrawing, setAutoDrawing] = useState(false);
  const [autoDrawQueue, setAutoDrawQueue] = useState([]);
  const [autoDrawHistory, setAutoDrawHistory] = useState([]);

  const [showAllWinners, setShowAllWinners] = useState(false);
  const [showCurrentPrizeWinners, setShowCurrentPrizeWinners] = useState(false);
  const [localConfetti, setLocalConfetti] = useState([]);
  const [localCongrats, setLocalCongrats] = useState([]);
  const [localBigWin, setLocalBigWin] = useState([]);

  const intervalRefs = useRef([null, null, null]);
  const [winnerPopup, setWinnerPopup] = useState({ open: false, winner: null });

  // Auto-draw queue REF (authoritative)
  const autoQueueRef = useRef([]);

  // Guard to prevent StrictMode double-run from duplicating the first winners
  const autoRunningRef = useRef(false);

  // Timers so we can clear on stop/unmount
  const timersRef = useRef({ spin: null, popup: null });

  const currentPrize = prizes[currentPrizeIndex];

  const pushWinner = useCallback((winnerObj) => {
    setPrizeWinners((prev) => ({
      ...prev,
      [winnerObj.prize]: [...(prev[winnerObj.prize] || []), winnerObj],
    }));
    setAllRoundWinners((prev) => [...prev, winnerObj]);
  }, []);

  // Previous winners for the popup list (exclude the current popup winner)
  const prevAutoWinners =
    autoDrawing && winnerPopup.open && winnerPopup.winner
      ? autoDrawHistory.filter((w) => w.timestamp !== winnerPopup.winner.timestamp)
      : [];

  useEffect(() => {
    setAllRoundWinners(initialWinners || []);
    setPrizeWinners(buildPrizeWinnersMap(initialWinners || []));
  }, [initialWinners]);

  // keep parent in sync (persisted)
  useEffect(() => {
    if (onWinnersChange) onWinnersChange(allRoundWinners);
  }, [allRoundWinners, onWinnersChange]);

  useEffect(() => {
    if (onPrizeIndexChange) onPrizeIndexChange(currentPrizeIndex);
  }, [currentPrizeIndex, onPrizeIndexChange]);

  // ✅ Spin audio must match VISUAL spinning only (no spin sound during popup)
  useEffect(() => {
    const isVisuallySpinning =
      allSpinning ||
      drawingState === "slot0spinning" ||
      drawingState === "slot1spinning" ||
      drawingState === "slot2spinning";

    if (isVisuallySpinning) audio.startSpin();
    else audio.stopSpin();
  }, [allSpinning, drawingState, audio]);

  const externalExcludedSet = useMemo(() => {
    return new Set((externalExcluded || []).map((w) => (w.numbers || []).join("-")));
  }, [externalExcluded]);

  const currentPrizeWinners = prizeWinners[currentPrize?.name] || [];
  const drawnCount = currentPrizeWinners.length;
  const remaining = currentPrize ? currentPrize.quantity - drawnCount : 0;

  const eligiblePool = useMemo(() => {
    const wonSet = new Set((allRoundWinners || []).map((w) => (w.numbers || []).join("-")));
    return (participants || []).filter((p) => {
      const key = (p.numbers || []).join("-");
      if (!key) return false;
      if (wonSet.has(key)) return false;
      if (externalExcludedSet.has(key)) return false;
      return true;
    });
  }, [participants, allRoundWinners, externalExcludedSet]);

  // Number spinning intervals (working logic)
  useEffect(() => {
    const source = eligiblePool.length > 0 ? eligiblePool : participants;

    if (!source || source.length === 0) {
      intervalRefs.current.forEach((it) => it && clearInterval(it));
      intervalRefs.current = [null, null, null];
      return;
    }

    if (allSpinning) {
      const intervals = [0, 1, 2].map((idx) =>
        setInterval(() => {
          setDisplayNumbers((prev) => {
            const next = [...prev];
            const rp = source[Math.floor(Math.random() * source.length)];
            next[idx] = rp.numbers[idx];
            return next;
          });
        }, 50)
      );
      intervalRefs.current = intervals;
    } else if (drawingState === "slot0spinning") {
      const interval = setInterval(() => {
        const rp = source[Math.floor(Math.random() * source.length)];
        setDisplayNumbers((prev) => [rp.numbers[0], prev[1], prev[2]]);
      }, 50);
      intervalRefs.current[0] = interval;
    } else if (drawingState === "slot1spinning") {
      if (intervalRefs.current[0]) clearInterval(intervalRefs.current[0]);
      const interval = setInterval(() => {
        const rp = source[Math.floor(Math.random() * source.length)];
        setDisplayNumbers((prev) => [prev[0], rp.numbers[1], prev[2]]);
      }, 50);
      intervalRefs.current[1] = interval;
    } else if (drawingState === "slot2spinning") {
      if (intervalRefs.current[1]) clearInterval(intervalRefs.current[1]);
      const interval = setInterval(() => {
        const rp = source[Math.floor(Math.random() * source.length)];
        setDisplayNumbers((prev) => [prev[0], prev[1], rp.numbers[2]]);
      }, 50);
      intervalRefs.current[2] = interval;
    } else {
      intervalRefs.current.forEach((it) => it && clearInterval(it));
      intervalRefs.current = [null, null, null];
    }

    return () => {
      intervalRefs.current.forEach((it) => it && clearInterval(it));
    };
  }, [drawingState, allSpinning, participants, eligiblePool]);

  const perWinnerMoney =
    isMoneyLike && moneyTotal && moneyWinnersCount
      ? Math.floor(parseInt(moneyTotal, 10) / parseInt(moneyWinnersCount, 10))
      : 0;

  const handleSpaceBar = useCallback(() => {
    audio.prime();
    audio.ensureBgmPlaying();

    if (!currentPrize) return;
    if (remaining <= 0) return;
    if (eligiblePool.length === 0) return;
    if (autoDrawing) return;
    if (winnerPopup.open) return;

    if (!allSpinning && !currentWinner) {
      const winner = eligiblePool[Math.floor(Math.random() * eligiblePool.length)];
      setCurrentWinner(winner);
      setDisplayNumbers([0, 0, 0]);
      setAllSpinning(true);
    } else if (allSpinning && currentWinner) {
      // ✅ stop sound EXACTLY when spinning stops & numbers appear
      audio.stopSpin();
      audio.stopTick(false);

      setDisplayNumbers([...currentWinner.numbers]);
      setAllSpinning(false);

      setTimeout(() => {
        const newWinner = {
          name: currentWinner.name,
          numbers: [...currentWinner.numbers],
          prize: currentPrize.name,
          timestamp: Date.now(),
        };

        setPrizeWinners((prev) => ({
          ...prev,
          [currentPrize.name]: [...(prev[currentPrize.name] || []), newWinner],
        }));
        setAllRoundWinners((prev) => [...prev, newWinner]);

        // (stop_tick is win now; no extra win sound)
        setCurrentWinner(null);
      }, 1000);
    }
  }, [audio, currentPrize, remaining, eligiblePool, allSpinning, currentWinner, autoDrawing, winnerPopup.open]);

  const handleOKey = useCallback(() => {
    audio.prime();
    audio.ensureBgmPlaying();

    if (!currentPrize) return;
    if (remaining <= 0) return;
    if (eligiblePool.length === 0) return;
    if (autoDrawing) return;
    if (winnerPopup.open) return;

    if (drawingState === "idle") {
      const winner = eligiblePool[Math.floor(Math.random() * eligiblePool.length)];
      setCurrentWinner(winner);
      setDisplayNumbers([0, 0, 0]);
      setDrawingState("slot0spinning");
    } else if (drawingState === "slot0spinning") {
      audio.stopTick(false);
      setDisplayNumbers((prev) => [currentWinner.numbers[0], prev[1], prev[2]]);
      setDrawingState("slot1spinning");
    } else if (drawingState === "slot1spinning") {
      audio.stopTick(false);
      setDisplayNumbers((prev) => [prev[0], currentWinner.numbers[1], prev[2]]);
      setDrawingState("slot2spinning");
    } else if (drawingState === "slot2spinning") {
      // final stop
      audio.stopSpin();
      audio.stopTick(false);

      setDisplayNumbers([...currentWinner.numbers]);
      setDrawingState("complete");

      setTimeout(() => {
        const newWinner = {
          name: currentWinner.name,
          numbers: [...currentWinner.numbers],
          prize: currentPrize.name,
          timestamp: Date.now(),
        };

        setPrizeWinners((prev) => ({
          ...prev,
          [currentPrize.name]: [...(prev[currentPrize.name] || []), newWinner],
        }));
        setAllRoundWinners((prev) => [...prev, newWinner]);

        setDrawingState("idle");
        setCurrentWinner(null);
      }, 1000);
    }
  }, [audio, currentPrize, remaining, eligiblePool, drawingState, currentWinner, autoDrawing, winnerPopup.open]);

  const handleDKey = useCallback(() => {
    if (remaining <= 0) return;
    if (autoDrawing) return;
    if (winnerPopup.open) return;
    if (eligiblePool.length === 0) return;

    // safety: don't start auto during manual spin/step mode
    if (allSpinning) return;
    if (drawingState !== "idle") return;
    if (currentWinner) return;

    const winnersNeeded = Math.min(remaining, eligiblePool.length);
    const selected = [];
    const temp = [...eligiblePool];

    for (let i = 0; i < winnersNeeded; i++) {
      const idx = Math.floor(Math.random() * temp.length);
      selected.push(temp[idx]);
      temp.splice(idx, 1);
    }

    autoQueueRef.current = selected.slice(); // authoritative
    setAutoDrawQueue(selected); // UI only
    setAutoDrawHistory([]); // reset history for this run
    setAutoDrawing(true);
  }, [remaining, autoDrawing, winnerPopup.open, eligiblePool, allSpinning, drawingState, currentWinner]);

  const handleRepeatRoll = useCallback(
    (winnerToReplace) => {
      audio.click();

      setUndoStack((prev) => [
        ...prev,
        {
          prizeWinners: JSON.parse(JSON.stringify(prizeWinners)),
          allRoundWinners: [...allRoundWinners],
          absentWinners: [...absentWinners],
        },
      ]);

      setAbsentWinners((prev) => [...prev, winnerToReplace.timestamp]);

      setPrizeWinners((prev) => {
        const updated = { ...prev };
        updated[winnerToReplace.prize] = (updated[winnerToReplace.prize] || []).filter((w) => w.timestamp !== winnerToReplace.timestamp);
        return updated;
      });

      setAllRoundWinners((prev) => prev.filter((w) => w.timestamp !== winnerToReplace.timestamp));
    },
    [audio, prizeWinners, allRoundWinners, absentWinners]
  );

  const handleUndo = useCallback(() => {
    if (undoStack.length === 0) return;
    audio.click();
    const last = undoStack[undoStack.length - 1];
    setPrizeWinners(last.prizeWinners);
    setAllRoundWinners(last.allRoundWinners);
    setAbsentWinners(last.absentWinners);
    setUndoStack((prev) => prev.slice(0, -1));
  }, [audio, undoStack]);

  const exportToExcel = useCallback(() => {
    audio.click();

    const exportData = [];

    prizes.forEach((prize) => {
      const winners = prizeWinners[prize.name] || [];
      const prizeValue = isMoneyLike ? `NT$ ${perWinnerMoney.toLocaleString()}` : prize.name;

      winners.forEach((winner) => {
        exportData.push({
          Prize: prizeValue,
          Winner: winner.name,
          Numbers: winner.numbers.join("-"),
        });
      });
    });

    const ws = XLSX.utils.json_to_sheet(exportData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, sheetName);
    XLSX.writeFile(wb, `Lucky_Draw_${sheetName}_Winners.xlsx`);
  }, [audio, prizes, prizeWinners, isMoneyLike, perWinnerMoney, sheetName]);

  // ✅ AUTO DRAW EFFECT
  useEffect(() => {
    if (!autoDrawing) {
      autoRunningRef.current = false;
      if (timersRef.current.spin) clearTimeout(timersRef.current.spin);
      if (timersRef.current.popup) clearTimeout(timersRef.current.popup);
      timersRef.current.spin = null;
      timersRef.current.popup = null;
      return;
    }

    if (autoRunningRef.current) return;
    autoRunningRef.current = true;

    let cancelled = false;

    const SPIN_MS = 1500;
    const POPUP_MS = 3000;

    const stopIntervals = () => {
      intervalRefs.current.forEach((it) => it && clearInterval(it));
      intervalRefs.current = [null, null, null];
    };

    const finishAuto = () => {
      autoRunningRef.current = false;
      setAutoDrawing(false);
      setAutoDrawQueue([]);
      setCurrentWinner(null);
      setAllSpinning(false);
      setWinnerPopup({ open: false, winner: null });
    };

    const runNext = () => {
      if (cancelled) return;

      // Periodic cleanup of audio pool during auto-draw
      audio.cleanupAudioPool();

      const prize = prizes[currentPrizeIndex];
      const q = autoQueueRef.current;

      // stop conditions
      if (!prize || !q || q.length === 0) {
        finishAuto();
        return;
      }

      // consume immediately (prevents duplicates)
      const winner = q.shift();
      setAutoDrawQueue(q.slice()); // UI reflection only

      // spin
      setCurrentWinner(winner);
      setDisplayNumbers([0, 0, 0]);
      setAllSpinning(true);

      timersRef.current.spin = setTimeout(() => {
        if (cancelled) return;

        stopIntervals();

        // ✅ stop spin + play stop tick EXACTLY at reveal
        setAllSpinning(false);
        setDisplayNumbers([...winner.numbers]);

        audio.stopSpin();
        audio.stopTick(true);

        const committed = {
          name: winner.name,
          numbers: [...winner.numbers],
          prize: prize.name,
          timestamp: Date.now() + Math.random(),
        };

        pushWinner(committed);
        setAutoDrawHistory((prev) => [...prev, committed]);

        setWinnerPopup({ open: true, winner: committed });

        timersRef.current.popup = setTimeout(() => {
          if (cancelled) return;

          setWinnerPopup({ open: false, winner: null });
          setCurrentWinner(null);

          runNext(); // next draw after popup delay
        }, POPUP_MS);
      }, SPIN_MS);
    };

    runNext();

    return () => {
      cancelled = true;
      autoRunningRef.current = false;
      if (timersRef.current.spin) clearTimeout(timersRef.current.spin);
      if (timersRef.current.popup) clearTimeout(timersRef.current.popup);
      timersRef.current.spin = null;
      timersRef.current.popup = null;
    };
  }, [autoDrawing, prizes, currentPrizeIndex, audio, pushWinner]);

  // Key handling (includes bonus key 6)
  useEffect(() => {
    const handleKeyPress = (e) => {
      audio.prime();
      audio.ensureBgmPlaying();

      const key = e.key;

      if (e.ctrlKey && (key === "z" || key === "Z")) {
        e.preventDefault();
        handleUndo();
        return;
      }

      if (key >= "0" && key <= "9") {
        e.preventDefault();
        const num = parseInt(key, 10);

        if (num === 2) {
          const newId = Date.now() + Math.random();
          setLocalBigWin((prev) => [...prev, newId]);
          setTimeout(() => setLocalBigWin((prev) => prev.filter((x) => x !== newId)), 6000);
        } else if (num === 4) {
          audio.confetti();
          const newId = Date.now() + Math.random();
          setLocalConfetti((prev) => [...prev, newId]);
          setTimeout(() => setLocalConfetti((prev) => prev.filter((x) => x !== newId)), 3000);
        } else if (num === 5) {
          const newId = Date.now() + Math.random();
          setLocalCongrats((prev) => [...prev, newId]);
          setTimeout(() => setLocalCongrats((prev) => prev.filter((x) => x !== newId)), 3000);
        } else if (num === 7) {
          // 7: NEXT ROUND (union -> general -> money -> bonus -> setup)
          audio.click();
          if (mode === "union") setCurrentPage("round2");
          else if (mode === "general") setCurrentPage("round3config");
          else if (mode === "money") setCurrentPage("bonusconfig");
          else if (mode === "bonus") setCurrentPage("setup");
        }

        return;
      }

      if (key === "ArrowLeft") {
        e.preventDefault();
        if (autoDrawing || winnerPopup.open) return;
        audio.click();
        setCurrentPrizeIndex((prev) => Math.max(0, prev - 1));
      } else if (key === "ArrowRight") {
        e.preventDefault();
        if (autoDrawing || winnerPopup.open) return;
        audio.click();
        setCurrentPrizeIndex((prev) => Math.min(prizes.length - 1, prev + 1));
      } else if (key === " ") {
        e.preventDefault();
        handleSpaceBar();
      } else if (key === "o" || key === "O") {
        e.preventDefault();
        handleOKey();
      } else if (key === "d" || key === "D") {
        e.preventDefault();
        handleDKey();
      } else if (key === "i" || key === "I") {
        e.preventDefault();
        audio.click();
        setShowAllWinners(true);
      } else if (key === "q" || key === "Q") {
        e.preventDefault();
        audio.click();
        setShowCurrentPrizeWinners(true);
      } else if (key === "Escape") {
        e.preventDefault();
        audio.click();
        setShowAllWinners(false);
        setShowCurrentPrizeWinners(false);
      }
    };

    window.addEventListener("keydown", handleKeyPress);
    return () => window.removeEventListener("keydown", handleKeyPress);
  }, [
    audio,
    mode,
    prizes.length,
    autoDrawing,
    winnerPopup.open,
    handleSpaceBar,
    handleOKey,
    handleDKey,
    handleUndo,
    exportToExcel,
    setCurrentPage,
  ]);

  if (!currentPrize || currentPrizeIndex >= prizes.length) {
    return (
      <div className="fixed inset-0 flex items-center justify-center bg-gradient-to-br from-red-700 via-red-600 to-yellow-600">
        <div className="w-full h-full flex items-center justify-center p-8" style={{ aspectRatio: "16/9", maxWidth: "100vw", maxHeight: "100vh" }}>
          <div className="bg-gradient-to-b from-yellow-50 to-red-50 rounded-3xl shadow-2xl p-12 text-center border-8 border-yellow-500">
            <div className="text-9xl mb-4">🏆</div>
            <h2 className="text-6xl font-bold text-red-700 mb-4">抽獎完成!</h2>
            <h3 className="text-4xl font-bold text-yellow-600 mb-6">All Prizes Drawn!</h3>

            {(mode === "union" || mode === "general" || mode === "money") && (
              <div className="flex flex-col gap-3">
                <button
                  onClick={() => {
                    audio.click();
                    if (mode === "union") setCurrentPage("round2");
                    else if (mode === "general") setCurrentPage("round3config");
                    else if (mode === "money") setCurrentPage("bonusconfig");
                  }}
                  className="bg-gradient-to-r from-purple-600 to-purple-700 text-white px-12 py-6 rounded-xl font-bold text-2xl hover:from-purple-700 hover:to-purple-800 transition border-4 border-yellow-400"
                >
                  下一輪 Next Round (7)
                </button>
              </div>
            )}

            {mode === "bonus" && (
              <div className="mt-4 text-2xl font-bold text-red-700">按 7 返回設定 Press 7 to return to Setup</div>
            )}

            <div className="mt-6">
              <button
                onClick={() => {
                  audio.click();
                  setCurrentPage("setup");
                }}
                className="bg-gradient-to-r from-gray-600 to-gray-700 text-white px-12 py-6 rounded-xl font-bold text-2xl hover:from-gray-700 hover:to-gray-800 transition border-4 border-yellow-400"
              >
                回設定 Back to Setup (9)
              </button>
            </div>
          </div>
        </div>
      </div>
    );
  }

  const Confetti = ({ id }) => {
    const confettiPieces = Array.from({ length: 150 }, (_, i) => ({
      id: i,
      left: Math.random() * 100,
      delay: Math.random() * 2,
      duration: 3 + Math.random() * 2,
      color: ["#FF0000", "#FFD700", "#FF6B6B", "#FFA500", "#FFFF00"][Math.floor(Math.random() * 5)],
    }));

    return (
      <div key={id} className="fixed inset-0 pointer-events-none z-50 animate-fadeOut">
        {confettiPieces.map((piece) => (
          <div
            key={piece.id}
            className="absolute w-4 h-4 animate-fall"
            style={{
              left: `${piece.left}%`,
              top: "-20px",
              backgroundColor: piece.color,
              animationDelay: `${piece.delay}s`,
              animationDuration: `${piece.duration}s`,
              transform: `rotate(${Math.random() * 360}deg)`,
            }}
          />
        ))}
        <style>{`
          @keyframes fall { to { transform: translateY(100vh) rotate(720deg); opacity: 0; } }
          @keyframes fadeOut { 0% { opacity: 1; } 70% { opacity: 1; } 100% { opacity: 0; } }
          .animate-fall { animation: fall linear forwards; }
          .animate-fadeOut { animation: fadeOut 3s forwards; }
        `}</style>
      </div>
    );
  };

  const CongratsPopup = ({ id }) => (
    <div key={id} className="fixed inset-0 pointer-events-none z-50 flex items-center justify-center animate-congratsFade">
      <div className="animate-congratsJump">
        <div
          className="font-bold text-yellow-300 drop-shadow-2xl leading-none"
          style={{
            fontFamily: "'KaiTi', 'STKaiti', 'BiauKai', 'DFKai-SB', serif",
            fontSize: "18rem",
            textShadow: "0 0 60px rgba(255, 215, 0, 1), 0 0 120px rgba(255, 215, 0, 0.9), 10px 10px 0 #FF0000, 20px 20px 0 #FF6B6B",
          }}
        >
          恭喜恭喜
        </div>
      </div>
      <style>{`
        @keyframes congratsJump {
          0% { transform: scale(0) translateY(100vh); opacity: 0; }
          50% { transform: scale(1.2) translateY(0); opacity: 1; }
          70% { transform: scale(0.9) translateY(-20px); }
          85% { transform: scale(1.05) translateY(0); }
          100% { transform: scale(1) translateY(0); opacity: 1; }
        }
        @keyframes congratsFade { 0% { opacity: 1; } 70% { opacity: 1; } 100% { opacity: 0; } }
        .animate-congratsJump { animation: congratsJump 1s ease-out forwards; }
        .animate-congratsFade { animation: congratsFade 3s forwards; }
      `}</style>
    </div>
  );

  const BigWin = ({ id }) => {
    const positions = [10, 30, 50, 70];
    return (
      <div key={id} className="fixed inset-0 pointer-events-none z-50 overflow-hidden">
        {positions.map((top, i) => (
          <div key={i} className="absolute animate-slideLR" style={{ top: `${top}%`, width: "100%" }}>
            <div
              className="font-bold text-red-600 whitespace-nowrap"
              style={{
                fontFamily: "'KaiTi', 'STKaiti', 'BiauKai', 'DFKai-SB', serif",
                fontSize: "10rem",
                textShadow: "5px 5px 0 #FFD700, 10px 10px 0 #FFA500",
              }}
            >
              大獎誕生
            </div>
          </div>
        ))}
        <style>{`
          @keyframes slideLR {
            0% { transform: translateX(-150%); opacity: 0; }
            15% { opacity: 1; }
            85% { opacity: 1; }
            100% { transform: translateX(150%); opacity: 0; }
          }
          .animate-slideLR { animation: slideLR 5s ease-in-out forwards; }
        `}</style>
      </div>
    );
  };

  return (
    <div
      className="fixed inset-0 flex items-center justify-center bg-gradient-to-br from-red-700 via-red-600 to-yellow-600"
      style={{ fontFamily: "'KaiTi', 'STKaiti', 'BiauKai', 'DFKai-SB', serif" }}
      onMouseDown={() => {
        audio.prime();
        audio.ensureBgmPlaying();
      }}
    >
      {localConfetti.map((id) => (
        <Confetti key={id} id={id} />
      ))}
      {localCongrats.map((id) => (
        <CongratsPopup key={id} id={id} />
      ))}
      {localBigWin.map((id) => (
        <BigWin key={id} id={id} />
      ))}

      {/* ALL WINNERS MODAL */}
      {showAllWinners && (
        <div
          className="fixed inset-0 bg-black bg-opacity-80 z-50 flex items-center justify-center p-2 md:p-6"
          onClick={() => setShowAllWinners(false)}
        >
          <div
            className="bg-gradient-to-b from-yellow-50 to-red-50 rounded-3xl shadow-2xl border-8 border-yellow-500 w-[min(1800px,98vw)] h-[94vh] flex flex-col"
            onClick={(e) => e.stopPropagation()}
          >
            <div className="text-center p-6 pb-3 flex-shrink-0">
              <h2 className="text-5xl font-bold text-red-700 mb-2">全部中獎名單 All Winners</h2>
              <p className="text-2xl text-yellow-600">總共 {allRoundWinners.length} 位中獎者</p>
            </div>

            <div className="flex-1 overflow-y-auto px-6 pb-3">
              {prizes.map((prize, idx) => {
                const winners = prizeWinners[prize.name] || [];
                if (winners.length === 0) return null;

                const title = isMoneyLike ? `NT$ ${perWinnerMoney.toLocaleString()}` : prize.name;

                return (
                  <div key={idx} className="mb-4 border-4 border-red-400 rounded-xl p-4 bg-white">
                    <h3 className="text-3xl font-bold text-red-800 mb-3">{title}</h3>
                    <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                      {winners.map((winner, wIdx) => (
                        <div
                          key={wIdx}
                          className="flex items-center gap-3 p-3 bg-gradient-to-r from-red-50 to-yellow-50 rounded-xl border-2 border-yellow-400"
                        >
                          <div className="flex gap-2">
                            {winner.numbers.map((num, numIdx) => (
                              <div
                                key={numIdx}
                                className="w-14 h-16 bg-gradient-to-b from-red-500 to-red-700 rounded-xl flex items-center justify-center border-4 border-yellow-400"
                              >
                                <span className="text-4xl font-extrabold text-yellow-300 leading-none">
                                  {num}
                                </span>
                              </div>
                            ))}
                          </div>

                          {/* Name unchanged size (only numbers got bigger) */}
                          <span className="font-bold text-red-800 text-lg">{winner.name}</span>
                        </div>
                      ))}
                    </div>
                  </div>
                );
              })}
            </div>

            <div className="p-6 pt-3 flex-shrink-0">
              <button
                onClick={() => {
                  audio.click();
                  setShowAllWinners(false);
                }}
                className="w-full bg-gradient-to-r from-red-600 to-red-700 text-white py-4 rounded-xl font-bold text-2xl hover:from-red-700 hover:to-red-800 transition border-4 border-yellow-400"
              >
                關閉 Close (ESC)
              </button>
            </div>
          </div>
        </div>
      )}

      {/* CURRENT PRIZE WINNERS MODAL */}
      {showCurrentPrizeWinners && (
        <div
          className="fixed inset-0 bg-black bg-opacity-80 z-50 flex items-center justify-center p-2 md:p-8"
          onClick={() => setShowCurrentPrizeWinners(false)}
        >
          <div
            className="bg-gradient-to-b from-yellow-50 to-red-50 rounded-3xl shadow-2xl border-8 border-yellow-500 w-[min(1600px,96vw)] max-h-[90vh] flex flex-col"
            onClick={(e) => e.stopPropagation()}
          >
            <div className="text-center p-8 pb-4">
              <h2 className="text-5xl font-bold text-red-700 mb-2">本獎項中獎名單 Current Prize Winners</h2>
              <h3 className="text-4xl font-bold text-yellow-600 mb-2">{isMoneyLike ? `NT$ ${perWinnerMoney.toLocaleString()}` : currentPrize.name}</h3>
              <p className="text-2xl text-red-600">共 {currentPrizeWinners.length} 位中獎者</p>
            </div>

            <div className="flex-1 overflow-y-auto px-8 pb-4">
              {currentPrizeWinners.length > 0 ? (
                <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                  {[...currentPrizeWinners].reverse().map((winner, idx) => (
                    <div key={idx} className="flex items-center gap-3 p-4 bg-gradient-to-r from-red-50 to-yellow-50 rounded-xl border-4 border-yellow-400">
                      <div className="flex gap-2">
                        {winner.numbers.map((num, numIdx) => (
                          <div
                            key={numIdx}
                            className="w-16 h-20 bg-gradient-to-b from-red-500 to-red-700 rounded-xl flex items-center justify-center border-4 border-yellow-400"
                          >
                            <span className="text-5xl font-extrabold text-yellow-300 leading-none">
                              {num}
                            </span>
                          </div>
                        ))}
                      </div>
                      <span className="font-bold text-red-800 text-2xl">{winner.name}</span>
                    </div>
                  ))}
                </div>
              ) : (
                <div className="text-center text-red-500 text-3xl font-bold p-8">尚無中獎者</div>
              )}
            </div>

            <div className="p-8 pt-4">
              <button
                onClick={() => {
                  audio.click();
                  setShowCurrentPrizeWinners(false);
                }}
                className="w-full bg-gradient-to-r from-red-600 to-red-700 text-white py-4 rounded-xl font-bold text-2xl hover:from-red-700 hover:to-red-800 transition border-4 border-yellow-400"
              >
                關閉 Close (ESC)
              </button>
            </div>
          </div>
        </div>
      )}

      {/* AUTO DRAW POPUP */}
      {winnerPopup.open && winnerPopup.winner && (
        <div className="fixed inset-0 z-[9999] flex items-center justify-center bg-black/70">
          <div className="bg-gradient-to-b from-yellow-50 to-red-50 rounded-3xl shadow-2xl border-8 border-yellow-500 p-10 w-[min(1400px,96vw)] max-h-[92vh] overflow-hidden">
            <div className="flex flex-col md:flex-row gap-10">
              {/* CURRENT WINNER (left) */}
              <div className="flex-1">
                <div className="text-center">
                  <div className="text-7xl font-extrabold text-red-700 mb-4 tracking-widest">WINNER</div>

                  {/* FW02 bigger */}
                  <div className="text-7xl font-extrabold text-red-800 mb-8 leading-none">
                    {winnerPopup.winner.name}
                  </div>

                  {/* 1 0 2 bigger */}
                  <div className="flex justify-center gap-6 mb-6 flex-wrap">
                    {winnerPopup.winner.numbers.map((n, i) => (
                      <div
                        key={i}
                        className="w-40 h-52 bg-gradient-to-b from-red-500 to-red-700 rounded-2xl shadow-2xl flex items-center justify-center border-6 border-yellow-400"
                      >
                        <span className="text-9xl font-extrabold text-yellow-300 leading-none drop-shadow-lg">
                          {n}
                        </span>
                      </div>
                    ))}
                  </div>

                  {/* PQB1 bigger */}
                  <div className="text-4xl font-extrabold text-yellow-700">
                    {winnerPopup.winner.prize}
                  </div>
                </div>
              </div>

              {/* PREVIOUS WINNERS (right) - ONLY DURING AUTO DRAW */}
              {autoDrawing && (
                <div className="w-full md:w-[420px] bg-white/90 rounded-2xl border-4 border-yellow-400 shadow-lg p-5">
                  <div className="flex items-baseline justify-between mb-2">
                    <div className="text-2xl font-bold text-red-700">Previous Winners</div>
                    <div className="text-2xl text-red-700 font-bold">{prevAutoWinners.length}</div>
                  </div>

                  <div className="text-base text-red-600 font-bold mb-3">
                    Auto-draw is running • Showing latest winners
                  </div>

                  <div className="max-h-[420px] overflow-y-auto space-y-3 pr-1">
                    {prevAutoWinners.length === 0 ? (
                      <div className="text-red-500 font-bold text-lg">No previous winners yet.</div>
                    ) : (
                      [...prevAutoWinners].slice(-10).reverse().map((w) => (
                        <div
                          key={w.timestamp}
                          className="flex items-center justify-between bg-gradient-to-r from-red-50 to-yellow-50 rounded-xl border-2 border-yellow-300 px-4 py-3"
                        >
                          <div className="font-bold text-red-800 truncate mr-3 text-xl">{w.name}</div>
                          <div className="text-red-700 font-bold whitespace-nowrap text-xl">
                            {w.numbers.join("-")}
                          </div>
                        </div>
                      ))
                    )}
                  </div>
                </div>
              )}
            </div>
          </div>
        </div>
      )}

      <div className="w-full h-full flex items-center justify-center p-4" style={{ aspectRatio: "16/9", maxWidth: "100vw", maxHeight: "100vh" }}>
        <div className="w-full h-full max-w-[177.78vh] max-h-[56.25vw]">
          <div className="bg-gradient-to-b from-yellow-50 to-red-50 rounded-3xl shadow-2xl p-6 border-8 border-yellow-500 h-full flex flex-col">
            <div className="text-center mb-2">
              <div className="flex items-center justify-center gap-6 mb-1">
                <span className="text-6xl">🎊</span>
                <h1 className="text-6xl font-bold text-red-700">{titleZh}</h1>
                <span className="text-6xl">🎊</span>
              </div>
              <div className="flex items-center justify-center gap-2 mb-2">
                <span className="text-2xl text-yellow-500">✦</span>
                <h2 className="text-3xl font-bold text-yellow-600">{titleEn}</h2>
                <span className="text-2xl text-yellow-500">✦</span>
              </div>

              {isMoneyLike && moneyTotal && moneyWinnersCount && (
                <div className="flex items-center justify-center gap-3 mb-1">
                  <span className="text-2xl text-yellow-600">💰</span>
                  <p className="text-3xl text-yellow-600 font-bold">每人獎金 Prize per Winner: NT$ {perWinnerMoney.toLocaleString()}</p>
                  <span className="text-2xl text-yellow-600">💰</span>
                </div>
              )}

              <div className="flex justify-center gap-3">
                <button
                  onClick={() => {
                    audio.click();
                    setShowAllWinners(true);
                  }}
                  className="bg-white px-5 py-2 rounded-xl border-4 border-yellow-400 font-bold text-red-700 hover:bg-yellow-50"
                >
                  查看全部得獎者 All Winners
                </button>
                <button
                  onClick={() => {
                    audio.click();
                    setShowCurrentPrizeWinners(true);
                  }}
                  className="bg-white px-5 py-2 rounded-xl border-4 border-red-400 font-bold text-red-700 hover:bg-red-50"
                >
                  查看本獎項得獎者 Current Prize
                </button>
              </div>
            </div>

            {!isMoneyLike && (
              <div className="flex items-center justify-center gap-3 mb-4">
                <button
                  onClick={() => {
                    audio.click();
                    setCurrentPrizeIndex(Math.max(0, currentPrizeIndex - 1));
                  }}
                  disabled={currentPrizeIndex === 0 || autoDrawing || winnerPopup.open}
                  className="bg-gradient-to-r from-gray-500 to-gray-600 text-white px-4 py-2 rounded-xl font-bold text-base hover:from-gray-600 hover:to-gray-700 transition disabled:opacity-30 disabled:cursor-not-allowed border-4 border-yellow-400 flex items-center gap-2"
                >
                  <ChevronLeft className="w-5 h-5" />
                  上一個
                </button>

                <span className="text-2xl font-bold text-red-700">
                  {currentPrizeIndex + 1} / {prizes.length}
                </span>

                <button
                  onClick={() => {
                    audio.click();
                    setCurrentPrizeIndex(Math.min(prizes.length - 1, currentPrizeIndex + 1));
                  }}
                  disabled={currentPrizeIndex === prizes.length - 1 || autoDrawing || winnerPopup.open}
                  className="bg-gradient-to-r from-gray-500 to-gray-600 text-white px-4 py-2 rounded-xl font-bold text-base hover:from-gray-600 hover:to-gray-700 transition disabled:opacity-30 disabled:cursor-not-allowed border-4 border-yellow-400 flex items-center gap-2"
                >
                  下一個
                  <ChevronRight className="w-5 h-5" />
                </button>
              </div>
            )}

            <div className="mb-4 p-4 bg-gradient-to-r from-red-100 to-yellow-100 rounded-2xl shadow-lg border-4 border-red-400">
              <div className="flex flex-col items-center text-center">
                {currentPrize.image && currentPrize.image.trim() !== "" && (
                  <img
                    src={currentPrize.image}
                    alt={currentPrize.name}
                    onError={(e) => {
                      e.currentTarget.style.display = "none";
                    }}
                    className="w-40 h-40 object-cover rounded-2xl shadow-2xl border-6 border-yellow-400 mb-3"
                  />
                )}

                <h2 className="text-4xl font-bold text-red-800 mb-2">
                  {isMoneyLike ? `NT$ ${perWinnerMoney.toLocaleString()}` : currentPrize.name}
                </h2>

                <div className="flex gap-6 text-2xl justify-center">
                  <span className="font-bold text-red-700">總數: {currentPrize.quantity}</span>
                  <span className="font-bold text-red-700">剩餘: {remaining}</span>
                  <span className="font-bold text-red-700">已抽: {drawnCount}</span>
                </div>

                {isMoneyLike && externalExcluded?.length > 0 && (
                  <div className="mt-2 text-lg font-bold text-red-700">已排除 Round 2 中獎者: {externalExcluded.length} 人</div>
                )}
              </div>
            </div>

            {(drawingState !== "idle" || allSpinning || currentWinner || autoDrawing) ? (
              <div className="flex flex-col items-center my-6">
                {/* ✅ Make the “102 below” smaller so it won’t overflow */}
                <div className="flex gap-3 justify-center mb-3">
                  {displayNumbers.map((num, idx) => (
                    <div
                      key={idx}
                      className={`w-28 h-36 bg-gradient-to-b from-red-500 to-red-700 rounded-xl shadow-2xl flex items-center justify-center border-6 border-yellow-400 transform transition-all ${
                        allSpinning || autoDrawing
                          ? "scale-110 rotate-2"
                          : (drawingState === "slot0spinning" && idx === 0) ||
                            (drawingState === "slot1spinning" && idx === 1) ||
                            (drawingState === "slot2spinning" && idx === 2)
                          ? "scale-110 rotate-2"
                          : "scale-100"
                      }`}
                    >
                      <span className="text-7xl font-bold text-yellow-300 drop-shadow-lg leading-none">{num}</span>
                    </div>
                  ))}
                </div>

                <div className="text-3xl font-bold text-red-700">
                  {autoDrawing
                    ? `自動抽獎中... Auto Drawing... (${autoDrawQueue.length} 剩餘)`
                    : drawingState === "complete" || (!allSpinning && currentWinner && drawingState === "idle")
                    ? "恭喜中獎 WINNER!"
                    : ""}
                </div>

                {(autoDrawing || autoDrawQueue.length > 0) && (
                  <div className="mt-4 w-full max-w-2xl bg-white/90 rounded-xl border-4 border-yellow-400 p-3 shadow-lg">
                    <div className="flex items-center justify-between mb-2">
                      <div className="text-lg font-bold text-red-700">已抽出名單 (Auto-Draw)</div>
                      <div className="text-red-700 font-bold">
                        {autoDrawHistory.length} / {autoDrawHistory.length + autoDrawQueue.length}
                      </div>
                    </div>

                    <div className="max-h-44 overflow-y-auto space-y-2">
                      {autoDrawHistory.length === 0 ? (
                        <div className="text-red-500 font-bold">尚無中獎者</div>
                      ) : (
                        [...autoDrawHistory].slice(-12).reverse().map((w) => (
                          <div key={w.timestamp} className="flex items-center justify-between bg-gradient-to-r from-red-50 to-yellow-50 rounded-lg border-2 border-yellow-300 px-3 py-2">
                            <div className="font-bold text-red-800">{w.name}</div>
                            <div className="text-red-700 font-bold">{w.numbers.join("-")}</div>
                          </div>
                        ))
                      )}
                    </div>
                  </div>
                )}
              </div>
            ) : (
              <div className="text-center my-4">
                <button
                  onClick={() => {
                    audio.click();
                    handleSpaceBar();
                  }}
                  disabled={remaining <= 0 || autoDrawing || winnerPopup.open}
                  className="bg-gradient-to-r from-red-600 to-red-800 text-yellow-300 px-12 py-5 rounded-2xl font-bold text-3xl hover:from-red-700 hover:to-red-900 transition disabled:opacity-50 disabled:cursor-not-allowed shadow-2xl transform hover:scale-105 border-6 border-yellow-400"
                >
                  開始抽獎 START
                </button>
              </div>
            )}

            <div className="flex-1 overflow-auto">
              <h3 className="text-3xl font-bold text-red-700 mb-3 text-center">
                中獎名單 ({currentPrizeWinners.length} / {currentPrize.quantity})
                {isMoneyLike && ` - NT$ ${perWinnerMoney.toLocaleString()} each`}
              </h3>

              <div className="bg-white rounded-xl overflow-hidden shadow-lg border-4 border-yellow-400">
                {currentPrizeWinners.length > 0 ? (
                  <table className="w-full">
                    <thead className="bg-gradient-to-r from-red-600 to-red-700 text-yellow-300">
                      <tr>
                        <th className="px-3 py-2 text-left font-bold text-lg">姓名</th>
                        <th className="px-3 py-2 text-left font-bold text-lg">幸運號碼</th>
                        <th className="px-3 py-2 text-center font-bold text-lg">操作</th>
                      </tr>
                    </thead>
                    <tbody>
                      {[...currentPrizeWinners].reverse().map((winner) => (
                        <tr key={`${winner.name}-${winner.timestamp}`} className="border-t-4 border-yellow-200 hover:bg-red-50 transition">
                          <td className="px-3 py-2 font-bold text-lg text-red-800">{winner.name}</td>
                          <td className="px-3 py-2">
                            <div className="flex gap-1">
                              {winner.numbers.map((num, numIdx) => (
                                <div key={numIdx} className="w-10 h-12 bg-gradient-to-b from-red-500 to-red-700 rounded-lg shadow-lg flex items-center justify-center border-3 border-yellow-400">
                                  <span className="text-2xl font-bold text-yellow-300 drop-shadow-lg">{num}</span>
                                </div>
                              ))}
                            </div>
                          </td>
                          <td className="px-3 py-2 text-center">
                            <button
                              onClick={() => handleRepeatRoll(winner)}
                              className="bg-gradient-to-r from-orange-500 to-orange-600 text-white px-3 py-2 rounded-xl font-bold hover:from-orange-600 hover:to-orange-700 transition text-base border-3 border-yellow-400"
                            >
                              重抽
                            </button>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                ) : (
                  <div className="p-8 text-center text-red-500 text-2xl font-bold">尚無中獎者 No winners yet.</div>
                )}
              </div>

              <div className="mt-4 flex justify-center gap-3">
                <button onClick={exportToExcel} className="bg-gradient-to-r from-green-600 to-green-700 text-white px-6 py-3 rounded-xl font-bold border-4 border-yellow-400">
                  匯出 Excel
                </button>

                <button
                  onClick={() => {
                    audio.click();
                    setCurrentPage("setup");
                  }}
                  className="bg-gradient-to-r from-gray-600 to-gray-700 text-white px-6 py-3 rounded-xl font-bold border-4 border-yellow-400"
                >
                  回設定
                </button>
              </div>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
};

/* =========================
   MAIN
   ========================= */

const LuckyDrawSystem = () => {
  const audio = useAudioManager();

  const [currentPage, setCurrentPage] = useState("setup");

  const [unionMembers, setUnionMembers] = useState([]);
  const [allWorkers, setAllWorkers] = useState([]);

  const [unionPrizes, setUnionPrizes] = useState([]);
  const [generalPrizes, setGeneralPrizes] = useState([]);

  // Board (money) draw config
  const [round3TotalMoney, setRound3TotalMoney] = useState("");
  const [round3Winners, setRound3Winners] = useState("");

  // Bonus draw config (key 6)
  const [bonusTotalMoney, setBonusTotalMoney] = useState("");
  const [bonusWinners, setBonusWinners] = useState("");

  // Persisted winners
  const [generalRoundWinners, setGeneralRoundWinners] = useState([]); // Round 2 winners
  const [moneyRoundWinners, setMoneyRoundWinners] = useState([]); // Board draw winners
  const [bonusRoundWinners, setBonusRoundWinners] = useState([]); // Bonus draw winners
  const [unionRoundWinners, setUnionRoundWinners] = useState([]); // Round 1 winners

  // Persisted prize index per page
  const [round1PrizeIndex, setRound1PrizeIndex] = useState(0);
  const [round2PrizeIndex, setRound2PrizeIndex] = useState(0);
  const [round3PrizeIndex, setRound3PrizeIndex] = useState(0);
  const [bonusPrizeIndex, setBonusPrizeIndex] = useState(0);

  const [hydrated, setHydrated] = useState(false);

  const parseExcelPaste = (pastedText) => {
    const rows = pastedText.trim().split("\n");
    const data = rows.slice(1).map((row) => row.split("\t"));
    return data;
  };

  const handlePasteWorkers = (e, type) => {
    audio.prime();
    audio.ensureBgmPlaying();

    const pastedText = e.clipboardData.getData("text");
    const data = parseExcelPaste(pastedText);

    const workers = data
      .map((cols) => ({
        name: cols[0] || "",
        numbers: [parseInt(cols[1], 10) || 0, parseInt(cols[2], 10) || 0, parseInt(cols[3], 10) || 0],
      }))
      .filter((w) => w.name);

    if (type === "union") setUnionMembers(workers);
    else setAllWorkers(workers);
  };

  const handlePastePrizes = (e, type) => {
    audio.prime();
    audio.ensureBgmPlaying();

    const pastedText = e.clipboardData.getData("text");
    const data = parseExcelPaste(pastedText);

    const prizes = data
      .map((cols) => ({
        name: (cols[0] || "").trim(),
        image: normalizeImageSrc(cols[1] || ""),
        quantity: parseInt(cols[2], 10) || 1,
      }))
      .filter((p) => p.name);

    if (type === "union") setUnionPrizes(prizes);
    else setGeneralPrizes(prizes);
  };

  // ✅ MUST be memoized so auto-draw effect doesn't restart from new array refs
  const round3PrizesMemo = useMemo(
    () => [
      {
        name: "董事會獎金 Board Prize",
        image: "",
        quantity: parseInt(round3Winners || "0", 10) || 0,
      },
    ],
    [round3Winners]
  );

  // ✅ MUST be memoized so auto-draw effect doesn't restart from new array refs
  const bonusPrizesMemo = useMemo(
    () => [
      {
        name: "加碼獎金 Bonus Prize",
        image: "",
        quantity: parseInt(bonusWinners || "0", 10) || 0,
      },
    ],
    [bonusWinners]
  );

  // Hydrate
  useEffect(() => {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (raw) {
      try {
        const saved = JSON.parse(raw);

        setCurrentPage(saved.currentPage ?? "setup");
        setUnionMembers(saved.unionMembers ?? []);
        setAllWorkers(saved.allWorkers ?? []);
        setUnionPrizes(saved.unionPrizes ?? []);
        setGeneralPrizes(saved.generalPrizes ?? []);

        setRound3TotalMoney(saved.round3TotalMoney ?? "");
        setRound3Winners(saved.round3Winners ?? "");

        setBonusTotalMoney(saved.bonusTotalMoney ?? "");
        setBonusWinners(saved.bonusWinners ?? "");

        setGeneralRoundWinners(saved.generalRoundWinners ?? []);
        setMoneyRoundWinners(saved.moneyRoundWinners ?? []);
        setBonusRoundWinners(saved.bonusRoundWinners ?? []);
        setUnionRoundWinners(saved.unionRoundWinners ?? []);

        setRound1PrizeIndex(saved.round1PrizeIndex ?? 0);
        setRound2PrizeIndex(saved.round2PrizeIndex ?? 0);
        setRound3PrizeIndex(saved.round3PrizeIndex ?? 0);
        setBonusPrizeIndex(saved.bonusPrizeIndex ?? 0);
      } catch {
        // ignore
      }
    }
    setHydrated(true);
  }, []);

  // Save
  useEffect(() => {
    if (!hydrated) return;
    const toSave = {
      currentPage,
      unionMembers,
      allWorkers,
      unionPrizes,
      generalPrizes,

      round3TotalMoney,
      round3Winners,

      bonusTotalMoney,
      bonusWinners,

      generalRoundWinners,
      moneyRoundWinners,
      bonusRoundWinners,
      unionRoundWinners,

      round1PrizeIndex,
      round2PrizeIndex,
      round3PrizeIndex,
      bonusPrizeIndex,
    };
    localStorage.setItem(STORAGE_KEY, JSON.stringify(toSave));
  }, [
    hydrated,
    currentPage,
    unionMembers,
    allWorkers,
    unionPrizes,
    generalPrizes,
    round3TotalMoney,
    round3Winners,
    bonusTotalMoney,
    bonusWinners,
    generalRoundWinners,
    moneyRoundWinners,
    bonusRoundWinners,
    unionRoundWinners,
    round1PrizeIndex,
    round2PrizeIndex,
    round3PrizeIndex,
    bonusPrizeIndex,
  ]);

  const validPages = new Set(["setup", "round1", "round2", "round3config", "round3", "bonusconfig", "bonus"]);

  if (!validPages.has(currentPage)) {
    return (
      <div className="fixed inset-0 flex items-center justify-center bg-black text-white">
        <div style={{ textAlign: "center" }}>
          <div style={{ fontSize: 24, fontWeight: 700 }}>Invalid currentPage</div>
          <div style={{ marginTop: 12 }}>currentPage = {String(currentPage)}</div>
          <button
            style={{ marginTop: 20, padding: "12px 18px", background: "#c00", color: "white", borderRadius: 10 }}
            onClick={() => {
              localStorage.removeItem(STORAGE_KEY);
              setCurrentPage("setup");
            }}
          >
            Reset to Setup (clears saved state)
          </button>
        </div>
      </div>
    );
  }

  const SetupPage = ({ audio }) => (
    <div
      className="fixed inset-0 flex items-center justify-center bg-gradient-to-br from-red-700 via-red-600 to-yellow-600"
      style={{ fontFamily: "'KaiTi', 'STKaiti', 'BiauKai', 'DFKai-SB', serif" }}
      onMouseDown={() => {
        audio.prime();
        audio.ensureBgmPlaying();
      }}
    >
      <div className="w-full h-full flex items-center justify-center p-8" style={{ aspectRatio: "16/9", maxWidth: "100vw", maxHeight: "100vh" }}>
        <div className="w-full h-full max-w-[177.78vh] max-h-[56.25vw]">
          <div className="bg-gradient-to-b from-red-50 to-yellow-50 rounded-3xl shadow-2xl p-8 border-8 border-yellow-500 h-full overflow-auto">
            <div className="text-center mb-6">
              <div className="text-6xl mb-3">🧧</div>
              <h1 className="text-5xl font-bold text-red-700 mb-2">新年抽獎系統</h1>
              <h2 className="text-3xl font-bold text-yellow-600 mb-2">Lucky Draw Event System</h2>
              <p className="text-red-600 text-xl">恭喜發財 • 萬事如意</p>
            </div>

            {/* Volume Control */}
            <div className="flex items-center justify-center mb-6 bg-gradient-to-r from-yellow-100 to-red-100 rounded-xl p-4 border-4 border-yellow-400 shadow-lg">
              <div className="flex items-center gap-4">
                <span className="text-2xl">🔊</span>
                <label className="text-lg font-bold text-red-700">背景音樂音量 BGM Volume:</label>
                <input
                  type="range"
                  min="0"
                  max="1"
                  step="0.01"
                  value={audio.bgmVolume}
                  onChange={(e) => audio.updateBgmVolume(parseFloat(e.target.value))}
                  className="w-32 h-3 bg-yellow-200 rounded-lg appearance-none cursor-pointer slider"
                  style={{
                    background: `linear-gradient(to right, #f59e0b 0%, #f59e0b ${audio.bgmVolume * 100}%, #fef3c7 ${audio.bgmVolume * 100}%, #fef3c7 100%)`,
                  }}
                />
                <span className="text-lg font-bold text-red-700 min-w-[3rem]">{Math.round(audio.bgmVolume * 100)}%</span>
              </div>
            </div>

            <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
              <div className="border-4 border-red-400 rounded-xl p-4 bg-gradient-to-br from-red-50 to-red-100 shadow-lg">
                <div className="flex items-center gap-2 mb-2">
                  <ClipboardPaste className="w-6 h-6 text-red-700" />
                  <h3 className="text-xl font-bold text-red-800">工會會員 | Union Members</h3>
                </div>
                <p className="text-sm text-red-700 mb-2">Copy from Excel: Name | Number1 | Number2 | Number3</p>
                <textarea
                  onPaste={(e) => handlePasteWorkers(e, "union")}
                  placeholder="Paste Excel data here..."
                  className="w-full h-24 p-2 border-4 border-red-300 rounded-lg focus:outline-none focus:border-yellow-500 text-sm bg-white"
                />
                {unionMembers.length > 0 && <div className="mt-2 text-green-700 font-bold text-base">✓ {unionMembers.length} members loaded</div>}
              </div>

              <div className="border-4 border-yellow-400 rounded-xl p-4 bg-gradient-to-br from-yellow-50 to-yellow-100 shadow-lg">
                <div className="flex items-center gap-2 mb-2">
                  <ClipboardPaste className="w-6 h-6 text-yellow-700" />
                  <h3 className="text-xl font-bold text-yellow-800">全體員工 | All Workers</h3>
                </div>
                <p className="text-sm text-yellow-700 mb-2">Copy from Excel: Name | Number1 | Number2 | Number3</p>
                <textarea
                  onPaste={(e) => handlePasteWorkers(e, "workers")}
                  placeholder="Paste Excel data here..."
                  className="w-full h-24 p-2 border-4 border-yellow-300 rounded-lg focus:outline-none focus:border-red-500 text-sm bg-white"
                />
                {allWorkers.length > 0 && <div className="mt-2 text-green-700 font-bold text-base">✓ {allWorkers.length} workers loaded</div>}
              </div>

              <div className="border-4 border-red-400 rounded-xl p-4 bg-gradient-to-br from-red-50 to-red-100 shadow-lg">
                <div className="flex items-center gap-2 mb-2">
                  <ClipboardPaste className="w-6 h-6 text-red-700" />
                  <h3 className="text-xl font-bold text-red-800">工會獎品 | Union Prizes</h3>
                </div>
                <p className="text-sm text-red-700 mb-2">Copy from Excel: Prize | Image URL or filename | Quantity</p>
                <textarea
                  onPaste={(e) => handlePastePrizes(e, "union")}
                  placeholder="Paste Excel data here..."
                  className="w-full h-24 p-2 border-4 border-red-300 rounded-lg focus:outline-none focus:border-yellow-500 text-sm bg-white"
                />
                {unionPrizes.length > 0 && <div className="mt-2 text-green-700 font-bold text-base">✓ {unionPrizes.length} prizes loaded</div>}
              </div>

              <div className="border-4 border-yellow-400 rounded-xl p-4 bg-gradient-to-br from-yellow-50 to-yellow-100 shadow-lg">
                <div className="flex items-center gap-2 mb-2">
                  <ClipboardPaste className="w-5 h-5 text-yellow-700" />
                  <h3 className="text-lg font-bold text-yellow-800">全員獎品 | General Prizes</h3>
                </div>
                <p className="text-xs text-yellow-700 mb-2">Copy from Excel: Prize | Image URL or filename | Quantity</p>
                <textarea
                  onPaste={(e) => handlePastePrizes(e, "general")}
                  placeholder="Paste Excel data here..."
                  className="w-full h-24 p-2 border-4 border-yellow-300 rounded-lg focus:outline-none focus:border-red-500 text-sm bg-white"
                />
                {generalPrizes.length > 0 && <div className="mt-2 text-green-700 font-bold text-base">✓ {generalPrizes.length} prizes loaded</div>}
              </div>
            </div>

            {unionMembers.length > 0 && allWorkers.length > 0 && unionPrizes.length > 0 && generalPrizes.length > 0 && (
              <div className="mt-6 flex gap-4">
                <button
                  onClick={() => {
                    audio.click();
                    audio.ensureBgmPlaying();
                    setCurrentPage("round1");
                  }}
                  className="flex-1 bg-gradient-to-r from-red-600 to-red-700 text-white py-4 rounded-xl font-bold text-xl hover:from-red-700 hover:to-red-800 transition shadow-lg border-4 border-yellow-400"
                  style={{ fontFamily: "'KaiTi', 'STKaiti', 'BiauKai', 'DFKai-SB', serif" }}
                >
                  🧧 福委會抽獎
                </button>
                <button
                  onClick={() => {
                    audio.click();
                    audio.ensureBgmPlaying();
                    setCurrentPage("round2");
                  }}
                  className="flex-1 bg-gradient-to-r from-yellow-500 to-yellow-600 text-red-800 py-4 rounded-xl font-bold text-xl hover:from-yellow-600 hover:to-yellow-700 transition shadow-lg border-4 border-red-400"
                  style={{ fontFamily: "'KaiTi', 'STKaiti', 'BiauKai', 'DFKai-SB', serif" }}
                >
                  🎊 廠商贊助 - 主管獎學金
                </button>
              </div>
            )}
          </div>
        </div>
      </div>
    </div>
  );

  const MoneyConfigPage = ({ titleZh, titleEn, totalMoney, setTotalMoney, winners, setWinners, onStart, onBackPage = "setup" }) => {
    const [tempMoney, setTempMoney] = useState(totalMoney || "");
    const [tempWinners, setTempWinners] = useState(winners || "");

    const prizePerWinner = tempMoney && tempWinners ? Math.floor(parseInt(tempMoney, 10) / parseInt(tempWinners, 10)) : 0;

    const handleStart = () => {
      if (tempMoney && tempWinners) {
        audio.click();
        setTotalMoney(tempMoney);
        setWinners(tempWinners);
        onStart();
      }
    };

    return (
      <div
        className="fixed inset-0 flex items-center justify-center bg-gradient-to-br from-red-700 via-red-600 to-yellow-600"
        style={{ fontFamily: "'KaiTi', 'STKaiti', 'BiauKai', 'DFKai-SB', serif" }}
        onMouseDown={() => {
          audio.prime();
          audio.ensureBgmPlaying();
        }}
      >
        <div className="w-full h-full flex items-center justify-center p-8" style={{ aspectRatio: "16/9", maxWidth: "100vw", maxHeight: "100vh" }}>
          <div className="w-full h-full max-w-[177.78vh] max-h-[56.25vw]">
            <div className="bg-gradient-to-b from-red-50 to-yellow-50 rounded-3xl shadow-2xl p-12 border-8 border-yellow-500 h-full flex flex-col items-center justify-center">
              <div className="text-center mb-8">
                <div className="text-6xl mb-4">💼</div>
                <h1 className="text-5xl font-bold text-red-700 mb-4">{titleZh}</h1>
                <h2 className="text-3xl font-bold text-yellow-600 mb-2">{titleEn}</h2>
              </div>

              <div className="w-full max-w-2xl space-y-6">
                <div className="border-4 border-red-400 rounded-xl p-6 bg-gradient-to-br from-red-50 to-red-100 shadow-lg">
                  <label className="block text-2xl font-bold text-red-800 mb-3">總獎金 Total Money (NT$):</label>
                  <input
                    type="number"
                    value={tempMoney}
                    onChange={(e) => setTempMoney(e.target.value)}
                    placeholder="例如: 120000"
                    className="w-full p-4 border-4 border-red-300 rounded-lg focus:outline-none focus:border-yellow-500 text-2xl bg-white"
                  />
                </div>

                <div className="border-4 border-yellow-400 rounded-xl p-6 bg-gradient-to-br from-yellow-50 to-yellow-100 shadow-lg">
                  <label className="block text-2xl font-bold text-yellow-800 mb-3">中獎人數 Number of Winners (1-12):</label>
                  <input
                    type="number"
                    min="1"
                    max="12"
                    value={tempWinners}
                    onChange={(e) => setTempWinners(e.target.value)}
                    placeholder="選擇 1-12"
                    className="w-full p-4 border-4 border-yellow-300 rounded-lg focus:outline-none focus:border-red-500 text-2xl bg-white"
                  />
                </div>

                {tempMoney && tempWinners && parseInt(tempWinners, 10) >= 1 && parseInt(tempWinners, 10) <= 12 && (
                  <div className="border-4 border-green-400 rounded-xl p-6 bg-gradient-to-br from-green-50 to-green-100 shadow-lg text-center">
                    <p className="text-2xl font-bold text-green-800">每人獎金 Prize per Winner:</p>
                    <p className="text-4xl font-bold text-green-600 mt-2">NT$ {prizePerWinner.toLocaleString()}</p>
                  </div>
                )}

                <button
                  onClick={handleStart}
                  disabled={!tempMoney || !tempWinners || parseInt(tempWinners, 10) < 1 || parseInt(tempWinners, 10) > 12}
                  className="w-full bg-gradient-to-r from-red-600 to-red-700 text-white py-6 rounded-xl font-bold text-3xl hover:from-red-700 hover:to-red-800 transition shadow-lg border-4 border-yellow-400 disabled:opacity-50 disabled:cursor-not-allowed"
                >
                  開始抽獎 Start Drawing
                </button>

                <button
                  onClick={() => {
                    audio.click();
                    setCurrentPage(onBackPage);
                  }}
                  className="w-full bg-gradient-to-r from-gray-500 to-gray-600 text-white py-4 rounded-xl font-bold text-xl hover:from-gray-600 hover:to-gray-700 transition shadow-lg border-4 border-gray-400 mt-4"
                >
                  返回 Back
                </button>
              </div>
            </div>
          </div>
        </div>
      </div>
    );
  };

  return (
    <div>
      {currentPage === "setup" && <SetupPage audio={audio} />}

      {currentPage === "round1" && (
        <DrawPage
          audio={audio}
          setCurrentPage={setCurrentPage}
          mode="union"
          titleZh="福委會抽獎"
          titleEn="Union Welfare Committee Draw"
          sheetName="福委會"
          prizes={unionPrizes}
          participants={unionMembers}
          initialWinners={unionRoundWinners}
          onWinnersChange={setUnionRoundWinners}
          externalExcluded={[]}
          initialPrizeIndex={round1PrizeIndex}
          onPrizeIndexChange={setRound1PrizeIndex}
        />
      )}

      {currentPage === "round2" && (
        <DrawPage
          audio={audio}
          setCurrentPage={setCurrentPage}
          mode="general"
          titleZh="廠商贊助 - 主管獎學金"
          titleEn="Vendor Sponsored / Management Scholarship"
          sheetName="廠商贊助"
          prizes={generalPrizes}
          participants={allWorkers}
          initialWinners={generalRoundWinners}
          onWinnersChange={setGeneralRoundWinners}
          externalExcluded={[]}
          initialPrizeIndex={round2PrizeIndex}
          onPrizeIndexChange={setRound2PrizeIndex}
        />
      )}

      {currentPage === "round3config" && (
        <MoneyConfigPage
          titleZh="董事會抽獎"
          titleEn="Board of Directors Lucky Draw"
          totalMoney={round3TotalMoney}
          setTotalMoney={setRound3TotalMoney}
          winners={round3Winners}
          setWinners={setRound3Winners}
          onStart={() => setCurrentPage("round3")}
          onBackPage="round2"
        />
      )}

      {currentPage === "round3" && (
        <DrawPage
          audio={audio}
          setCurrentPage={setCurrentPage}
          mode="money"
          titleZh="董事會抽獎"
          titleEn="Board of Directors Lucky Draw"
          sheetName="董事會"
          prizes={round3PrizesMemo}
          participants={allWorkers}
          initialWinners={moneyRoundWinners}
          onWinnersChange={setMoneyRoundWinners}
          externalExcluded={generalRoundWinners}
          initialPrizeIndex={round3PrizeIndex}
          onPrizeIndexChange={setRound3PrizeIndex}
          moneyTotal={round3TotalMoney}
          moneyWinnersCount={round3Winners}
        />
      )}

      {currentPage === "bonusconfig" && (
        <MoneyConfigPage
          titleZh="加碼抽獎"
          titleEn="Bonus Draw"
          totalMoney={bonusTotalMoney}
          setTotalMoney={setBonusTotalMoney}
          winners={bonusWinners}
          setWinners={setBonusWinners}
          onStart={() => setCurrentPage("bonus")}
          onBackPage="round2"
        />
      )}

      {currentPage === "bonus" && (
        <DrawPage
          audio={audio}
          setCurrentPage={setCurrentPage}
          mode="bonus"
          titleZh="加碼抽獎"
          titleEn="Bonus Draw"
          sheetName="加碼"
          prizes={bonusPrizesMemo}
          participants={allWorkers}
          initialWinners={bonusRoundWinners}
          onWinnersChange={setBonusRoundWinners}
          externalExcluded={generalRoundWinners}
          initialPrizeIndex={bonusPrizeIndex}
          onPrizeIndexChange={setBonusPrizeIndex}
          moneyTotal={bonusTotalMoney}
          moneyWinnersCount={bonusWinners}
        />
      )}
    </div>
  );
};

export default LuckyDrawSystem;
