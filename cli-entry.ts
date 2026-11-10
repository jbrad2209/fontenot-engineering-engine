import { runCli } from "./cli.ts";

runCli(process.argv.slice(2)).then(
  (code) => {
    process.exit(code);
  },
  (err: unknown) => {
    process.stderr.write((err instanceof Error ? err.message : String(err)) + "\n");
    process.exit(1);
  },
);
