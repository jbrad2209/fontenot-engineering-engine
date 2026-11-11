/** New viewer accounts must be a real IHG address. The existing admin is exempt. */
export function ihgEmail(value: string): string | null {
  const email = value.trim().toLowerCase();
  if (!/^[a-z0-9._%+-]+@ihg\.com$/.test(email)) return null;
  return email;
}
