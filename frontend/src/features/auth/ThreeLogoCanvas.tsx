import { useEffect, useRef, useState } from "react";
import * as THREE from "three";
import { GLTFLoader } from "three/examples/jsm/loaders/GLTFLoader.js";
import { DRACOLoader } from "three/examples/jsm/loaders/DRACOLoader.js";
import { LogoFallback } from "./LogoFallback";

interface ThreeLogoCanvasProps {
  onError?: () => void;
}

export default function ThreeLogoCanvas({ onError }: ThreeLogoCanvasProps) {
  const containerRef = useRef<HTMLDivElement>(null);
  const [hasError, setHasError] = useState(false);
  const [isLoading, setIsLoading] = useState(true);

  useEffect(() => {
    const container = containerRef.current;
    if (!container) return;

    // Inicialização do WebGLRenderer com parâmetros de alta fidelidade
    let renderer: THREE.WebGLRenderer;
    try {
      renderer = new THREE.WebGLRenderer({
        antialias: true,
        alpha: true,
        powerPreference: "high-performance",
      });
    } catch {
      setHasError(true);
      onError?.();
      return;
    }

    renderer.setPixelRatio(Math.min(window.devicePixelRatio, 2));
    renderer.setSize(container.clientWidth, container.clientHeight);
    renderer.toneMapping = THREE.ACESFilmicToneMapping;
    renderer.toneMappingExposure = 1.25;
    container.appendChild(renderer.domElement);

    const scene = new THREE.Scene();

    const camera = new THREE.PerspectiveCamera(
      38,
      container.clientWidth / container.clientHeight,
      0.1,
      100
    );
    // Modelo normalizado de tamanho ~2.0, posicionado em z=4.5 para enquadramento perfeito
    camera.position.set(0, 0, 4.5);

    // SISTEMA DE ILUMINAÇÃO DINÂMICO (Spotify + Solvd Navy)
    // 1. Luz ambiente rica em azul marinho
    const ambientLight = new THREE.AmbientLight(0x1a3158, 3.2);
    scene.add(ambientLight);

    // 2. Luz frontal para iluminar as letras "revalle" em relevo
    const frontLight = new THREE.DirectionalLight(0xe0f2fe, 3.5);
    frontLight.position.set(0, 1.5, 4.5);
    scene.add(frontLight);

    // 3. Ponto de luz orbital Ciano Elétrico (varredura contínua de reflexos)
    const cyanLight = new THREE.PointLight(0x38bdf8, 7.0, 15);
    cyanLight.position.set(3, 3, 3);
    scene.add(cyanLight);

    // 4. Luz de preenchimento inferior oposta
    const fillLight = new THREE.DirectionalLight(0x2563eb, 3.0);
    fillLight.position.set(-3, -3, 2);
    scene.add(fillLight);

    // 5. Backlight para silhueta 3D
    const backLight = new THREE.DirectionalLight(0x93c5fd, 2.0);
    backLight.position.set(0, 3, -3);
    scene.add(backLight);

    // Carregamento com decodificador DRACO local (100% self-hosted)
    const dracoLoader = new DRACOLoader();
    dracoLoader.setDecoderPath("/draco/");

    const gltfLoader = new GLTFLoader();
    gltfLoader.setDRACOLoader(dracoLoader);

    let modelGroup: THREE.Group | null = null;
    let animId: number;

    // Estado interativo de rotação e mouse
    let mouseParallaxX = 0;
    let mouseParallaxY = 0;
    let currentParallaxX = 0;
    let currentParallaxY = 0;

    // Suporte a rotação por arrasto (drag)
    let isDragging = false;
    let previousPointerX = 0;
    let previousPointerY = 0;
    let dragVelocityX = 0;
    let dragVelocityY = 0;
    let dragRotX = 0;
    let dragRotY = 0;

    gltfLoader.load(
      "/3d/logo.glb",
      (gltf) => {
        modelGroup = gltf.scene;

        // Aplicação do material azul marinho nobre com acabamento acetinado
        modelGroup.traverse((child) => {
          if ((child as THREE.Mesh).isMesh) {
            const mesh = child as THREE.Mesh;
            mesh.castShadow = true;
            mesh.receiveShadow = true;
            if (mesh.geometry) {
              mesh.geometry.computeVertexNormals();
            }
            mesh.material = new THREE.MeshStandardMaterial({
              color: 0x19396e, // Azul marinho nobre luminoso (não preto)
              metalness: 0.65,
              roughness: 0.28,
              flatShading: false,
            });
          }
        });

        modelGroup.position.set(0, 0, 0);
        scene.add(modelGroup);
        setIsLoading(false);

        // Loop de animação contínuo garantido
        const clock = new THREE.Clock();

        const animate = () => {
          animId = requestAnimationFrame(animate);
          const elapsed = clock.getElapsedTime();

          if (modelGroup) {
            // Suavização do mouse parallax
            currentParallaxX += (mouseParallaxX - currentParallaxX) * 0.05;
            currentParallaxY += (mouseParallaxY - currentParallaxY) * 0.05;

            // Inércia de arrasto
            if (!isDragging) {
              dragRotX += dragVelocityX;
              dragRotY += dragVelocityY;
              dragVelocityX *= 0.92;
              dragVelocityY *= 0.92;
              // Retorno suave à posição base
              dragRotX *= 0.96;
              dragRotY *= 0.96;
            }

            // Movimento Contínuo Elegante:
            // 1. Oscilação lenta em Y (sweep orbital de ~±24 graus mantendo o logo legível)
            const autoRotY = Math.sin(elapsed * 0.8) * 0.42;

            // 2. Leve inclinação tridimensional em X
            const autoRotX = 0.08 + Math.cos(elapsed * 0.65) * 0.12;

            // 3. Flutuação vertical de respiração
            const autoFloatY = Math.sin(elapsed * 1.3) * 0.12;

            // 4. Sutil roll em Z
            const autoTiltZ = Math.sin(elapsed * 0.9) * 0.05;

            // Aplicação das transformações combinadas
            modelGroup.position.y = autoFloatY;
            modelGroup.rotation.y = autoRotY + currentParallaxX + dragRotY;
            modelGroup.rotation.x = autoRotX - currentParallaxY + dragRotX;
            modelGroup.rotation.z = autoTiltZ;

            // Dinâmica de iluminação: ponto de luz ciano viaja suavemente criando reflexos no relevo
            cyanLight.position.x = 2.5 + Math.sin(elapsed * 1.1) * 2.2;
            cyanLight.position.y = 2.0 + Math.cos(elapsed * 0.9) * 1.8;
          }

          renderer.render(scene, camera);
        };

        animate();
      },
      undefined,
      (error) => {
        console.warn("Falha ao carregar modelo 3D local, ativando fallback vetorial:", error);
        setHasError(true);
        onError?.();
      }
    );

    // 1. Parallax global pelo movimento do mouse na tela inteira
    const handleGlobalPointerMove = (e: PointerEvent) => {
      const nx = (e.clientX / window.innerWidth) * 2 - 1; // [-1, 1]
      const ny = (e.clientY / window.innerHeight) * 2 - 1; // [-1, 1]

      mouseParallaxX = nx * 0.45;
      mouseParallaxY = ny * 0.35;
    };

    // 2. Interatividade por arrasto no canvas (Drag to Rotate)
    const handlePointerDown = (e: PointerEvent) => {
      isDragging = true;
      previousPointerX = e.clientX;
      previousPointerY = e.clientY;
      dragVelocityX = 0;
      dragVelocityY = 0;
    };

    const handleCanvasPointerMove = (e: PointerEvent) => {
      if (!isDragging) return;
      const deltaX = e.clientX - previousPointerX;
      const deltaY = e.clientY - previousPointerY;

      dragVelocityY = deltaX * 0.008;
      dragVelocityX = deltaY * 0.008;

      dragRotY += dragVelocityY;
      dragRotX += dragVelocityX;

      previousPointerX = e.clientX;
      previousPointerY = e.clientY;
    };

    const handlePointerUp = () => {
      isDragging = false;
    };

    window.addEventListener("pointermove", handleGlobalPointerMove, { passive: true });
    window.addEventListener("pointerup", handlePointerUp, { passive: true });

    container.addEventListener("pointerdown", handlePointerDown);
    container.addEventListener("pointermove", handleCanvasPointerMove);

    // Responsividade do canvas ao redimensionar
    const handleResize = () => {
      if (!container) return;
      const width = container.clientWidth;
      const height = container.clientHeight;
      camera.aspect = width / height;
      camera.updateProjectionMatrix();
      renderer.setSize(width, height);
    };

    const resizeObserver = new ResizeObserver(handleResize);
    resizeObserver.observe(container);

    // Limpeza de recursos WebGL ao desmontar
    return () => {
      if (animId) cancelAnimationFrame(animId);
      window.removeEventListener("pointermove", handleGlobalPointerMove);
      window.removeEventListener("pointerup", handlePointerUp);
      container.removeEventListener("pointerdown", handlePointerDown);
      container.removeEventListener("pointermove", handleCanvasPointerMove);
      resizeObserver.disconnect();

      dracoLoader.dispose();

      if (modelGroup) {
        modelGroup.traverse((child) => {
          if ((child as THREE.Mesh).isMesh) {
            const mesh = child as THREE.Mesh;
            mesh.geometry?.dispose();
            if (Array.isArray(mesh.material)) {
              mesh.material.forEach((m) => m.dispose());
            } else {
              mesh.material?.dispose();
            }
          }
        });
      }

      renderer.dispose();
      if (renderer.domElement && container.contains(renderer.domElement)) {
        container.removeChild(renderer.domElement);
      }
    };
  }, [onError]);

  if (hasError) {
    return <LogoFallback />;
  }

  return (
    <div className="relative size-full flex items-center justify-center select-none">
      {isLoading && (
        <div className="absolute inset-0 flex items-center justify-center z-10 transition-opacity duration-300">
          <LogoFallback />
        </div>
      )}
      <div
        ref={containerRef}
        className="size-full flex items-center justify-center cursor-grab active:cursor-grabbing touch-none"
        title="Arraste para girar o modelo 3D ou mova o mouse para efeito parallax"
        aria-label="Visualização tridimensional interativa do logo institucional"
      />
    </div>
  );
}
