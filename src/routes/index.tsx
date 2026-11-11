import { createFileRoute } from "@tanstack/react-router";
import { EngineApp } from "@/components/engine/engine-app";

export const Route = createFileRoute("/")({ component: Home });

function Home() {
  return <EngineApp />;
}
