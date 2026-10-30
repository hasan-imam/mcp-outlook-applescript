import { DELIMITERS } from './scripts.js';
/**
 * AppleScript that enumerates all Outlook accounts (Exchange, IMAP, POP)
 * and returns their ID, name, email, and type as delimited output.
 */
export const LIST_ACCOUNTS = `
tell application "Microsoft Outlook"
  set output to ""

  -- Enumerate Exchange accounts
  try
    set exchangeAccounts to every exchange account
    repeat with acc in exchangeAccounts
      try
        set accId to id of acc
        set accName to name of acc
        set accEmail to email address of acc
        set accType to "exchange"
        set output to output & "${DELIMITERS.RECORD}id${DELIMITERS.EQUALS}" & accId & "${DELIMITERS.FIELD}name${DELIMITERS.EQUALS}" & accName & "${DELIMITERS.FIELD}email${DELIMITERS.EQUALS}" & accEmail & "${DELIMITERS.FIELD}type${DELIMITERS.EQUALS}" & accType
      end try
    end repeat
  end try

  -- Enumerate IMAP accounts
  try
    set imapAccounts to every imap account
    repeat with acc in imapAccounts
      try
        set accId to id of acc
        set accName to name of acc
        set accEmail to email address of acc
        set accType to "imap"
        set output to output & "${DELIMITERS.RECORD}id${DELIMITERS.EQUALS}" & accId & "${DELIMITERS.FIELD}name${DELIMITERS.EQUALS}" & accName & "${DELIMITERS.FIELD}email${DELIMITERS.EQUALS}" & accEmail & "${DELIMITERS.FIELD}type${DELIMITERS.EQUALS}" & accType
      end try
    end repeat
  end try

  -- Enumerate POP accounts
  try
    set popAccounts to every pop account
    repeat with acc in popAccounts
      try
        set accId to id of acc
        set accName to name of acc
        set accEmail to email address of acc
        set accType to "pop"
        set output to output & "${DELIMITERS.RECORD}id${DELIMITERS.EQUALS}" & accId & "${DELIMITERS.FIELD}name${DELIMITERS.EQUALS}" & accName & "${DELIMITERS.FIELD}email${DELIMITERS.EQUALS}" & accEmail & "${DELIMITERS.FIELD}type${DELIMITERS.EQUALS}" & accType
      end try
    end repeat
  end try

  return output
end tell
`;
/**
 * AppleScript that resolves the default Outlook account.
 * Falls back through Exchange, IMAP, and POP if no explicit default is set.
 */
export const GET_DEFAULT_ACCOUNT = `
tell application "Microsoft Outlook"
  try
    -- Attempt to read the configured default account
    set defaultAcc to default account
    set accId to id of defaultAcc
    return "id${DELIMITERS.EQUALS}" & accId
  on error
    -- Fall back to the first Exchange account, then IMAP, then POP
    try
      set firstAcc to first exchange account
      set accId to id of firstAcc
      return "id${DELIMITERS.EQUALS}" & accId
    on error
      try
        set firstAcc to first imap account
        set accId to id of firstAcc
        return "id${DELIMITERS.EQUALS}" & accId
      on error
        try
          set firstAcc to first pop account
          set accId to id of firstAcc
          return "id${DELIMITERS.EQUALS}" & accId
        on error
          return "error${DELIMITERS.EQUALS}No accounts found"
        end try
      end try
    end try
  end try
end tell
`;
/**
 * Builds an AppleScript that lists mail folders for the given account IDs.
 * Iterates across Exchange, IMAP, and POP account types, returning each
 * folder's ID, name, unread count, message count, and parent account ID.
 * @param accountIds - Outlook account IDs whose folders to enumerate.
 * @returns A ready-to-execute AppleScript string.
 */
export function listMailFoldersByAccounts(accountIds) {
    const accountFilter = accountIds.map(id => `id = ${id}`).join(' or ');
    return `
tell application "Microsoft Outlook"
  set output to ""

  -- Collect folders from matching Exchange accounts
  try
    set targetAccounts to (every exchange account whose ${accountFilter})
    repeat with acc in targetAccounts
      set accId to id of acc
      set allFolders to mail folders of acc
      repeat with f in allFolders
        try
          set fId to id of f
          set fName to name of f
          set uCount to unread count of f
          set mCount to 0
          try
            set mCount to count of messages of f
          end try
          set output to output & "${DELIMITERS.RECORD}id${DELIMITERS.EQUALS}" & fId & "${DELIMITERS.FIELD}name${DELIMITERS.EQUALS}" & fName & "${DELIMITERS.FIELD}unreadCount${DELIMITERS.EQUALS}" & uCount & "${DELIMITERS.FIELD}messageCount${DELIMITERS.EQUALS}" & mCount & "${DELIMITERS.FIELD}accountId${DELIMITERS.EQUALS}" & accId
        end try
      end repeat
    end repeat
  end try

  -- Collect folders from matching IMAP accounts
  try
    set targetAccounts to (every imap account whose ${accountFilter})
    repeat with acc in targetAccounts
      set accId to id of acc
      set allFolders to mail folders of acc
      repeat with f in allFolders
        try
          set fId to id of f
          set fName to name of f
          set uCount to unread count of f
          set mCount to 0
          try
            set mCount to count of messages of f
          end try
          set output to output & "${DELIMITERS.RECORD}id${DELIMITERS.EQUALS}" & fId & "${DELIMITERS.FIELD}name${DELIMITERS.EQUALS}" & fName & "${DELIMITERS.FIELD}unreadCount${DELIMITERS.EQUALS}" & uCount & "${DELIMITERS.FIELD}messageCount${DELIMITERS.EQUALS}" & mCount & "${DELIMITERS.FIELD}accountId${DELIMITERS.EQUALS}" & accId
        end try
      end repeat
    end repeat
  end try

  -- Collect folders from matching POP accounts
  try
    set targetAccounts to (every pop account whose ${accountFilter})
    repeat with acc in targetAccounts
      set accId to id of acc
      set allFolders to mail folders of acc
      repeat with f in allFolders
        try
          set fId to id of f
          set fName to name of f
          set uCount to unread count of f
          set mCount to 0
          try
            set mCount to count of messages of f
          end try
          set output to output & "${DELIMITERS.RECORD}id${DELIMITERS.EQUALS}" & fId & "${DELIMITERS.FIELD}name${DELIMITERS.EQUALS}" & fName & "${DELIMITERS.FIELD}unreadCount${DELIMITERS.EQUALS}" & uCount & "${DELIMITERS.FIELD}messageCount${DELIMITERS.EQUALS}" & mCount & "${DELIMITERS.FIELD}accountId${DELIMITERS.EQUALS}" & accId
        end try
      end repeat
    end repeat
  end try

  return output
end tell
`;
}
