import { createFileRoute } from "@tanstack/react-router";

import { OffboardingHubPage } from "@/features/offboarding/OffboardingHubPage";

export const Route = createFileRoute("/_shell/offboarding-ti")({
  component: OffboardingHubPage,
});
