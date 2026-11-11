import { createFileRoute } from "@tanstack/react-router";
import { AdminPage } from "@/components/engine/admin-page";

export const Route = createFileRoute("/login")({ component: AdminPage });
