import { getContext } from '@microsoft/power-apps/app'

export type CurrentUser = {
  name: string
  email: string
}

export async function getCurrentUser(): Promise<CurrentUser> {
  const context = await getContext()
  return {
    name: context.user.fullName || context.user.userPrincipalName || '',
    email: (context.user.userPrincipalName || '').toLowerCase(),
  }
}

export function includesCurrentUser(value: unknown, email: string) {
  if (!email) return false
  const people = Array.isArray(value) ? value : value ? [value] : []
  return people.some((person) => {
    if (!person || typeof person !== 'object') return false
    const personEmail = (person as { Email?: unknown }).Email
    return typeof personEmail === 'string' && personEmail.toLowerCase() === email
  })
}