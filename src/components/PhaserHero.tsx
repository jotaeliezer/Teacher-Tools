import { useEffect, useRef } from "react";

const COLORS = [0xc1d82f, 0xffdd00, 0xaf006f, 0x111827];

export default function PhaserHero() {
  const containerRef = useRef<HTMLDivElement | null>(null);
  const gameRef = useRef<any>(null);

  useEffect(() => {
    let isMounted = true;

    async function start() {
      const container = containerRef.current;
      if (!container || gameRef.current) return;

      const PhaserModule = await import("phaser");
      if (!isMounted) return;
      const Phaser = PhaserModule.default;

      const config = {
        type: Phaser.AUTO,
        parent: container,
        width: container.clientWidth,
        height: container.clientHeight,
        transparent: true,
        physics: {
          default: "arcade",
          arcade: { gravity: { y: 0 }, debug: false },
        },
        scene: {
          create() {
            const { width, height } = this.scale;
            this.dots = [];
            for (let i = 0; i < 24; i += 1) {
              const radius = Phaser.Math.Between(14, 36);
              const dot = this.add.circle(
                Phaser.Math.Between(0, width),
                Phaser.Math.Between(0, height),
                radius,
                COLORS[i % COLORS.length],
                0.15
              );
              dot.vx = Phaser.Math.FloatBetween(-0.15, 0.15);
              dot.vy = Phaser.Math.FloatBetween(-0.12, 0.12);
              this.dots.push(dot);
            }
          },
          update(_, delta) {
            const { width, height } = this.scale;
            const step = delta * 0.08;
            this.dots.forEach((dot) => {
              dot.x += dot.vx * step;
              dot.y += dot.vy * step;
              if (dot.x < -40 || dot.x > width + 40) dot.vx *= -1;
              if (dot.y < -40 || dot.y > height + 40) dot.vy *= -1;
            });
          },
        },
      };

      const game = new Phaser.Game(config);
      gameRef.current = game;

      const handleResize = () => {
        if (!containerRef.current || !gameRef.current) return;
        gameRef.current.scale.resize(containerRef.current.clientWidth, containerRef.current.clientHeight);
      };

      window.addEventListener("resize", handleResize);

      game.events.once("destroy", () => {
        window.removeEventListener("resize", handleResize);
      });
    }

    start();

    return () => {
      isMounted = false;
      if (gameRef.current) {
        gameRef.current.destroy(true);
        gameRef.current = null;
      }
    };
  }, []);

  return <div className="phaser-hero" ref={containerRef} aria-hidden="true" />;
}
