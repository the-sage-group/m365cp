import { z } from "zod";

const envSchema = z.object({
  PORT: z.coerce.number().default(3001),
  ANTHROPIC_API_KEY: z.string().min(1, "ANTHROPIC_API_KEY is required"),
});

export const env = envSchema.parse(process.env);
