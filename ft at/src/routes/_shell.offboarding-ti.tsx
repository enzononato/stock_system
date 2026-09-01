import { createFileRoute } from "@tanstack/react-router";

import { TiOffboardingPage } from "@/features/offboarding/TiOffboardingPage";

export const Route = createFileRoute("/_shell/offboarding-ti")({
  component: TiOffboardingPage,
});
