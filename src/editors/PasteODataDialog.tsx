// Dialog that lets the user paste a Dataverse Web API URL / query string and
// have DRS parse-validate-apply it to the builder state.
//
// Three-phase pipeline:
//   1. Syntactic parse — `odataParser.parseODataUrl` lexes and builds the AST.
//   2. Static validation — entity set known, query options known, DV functions
//      in the operator registry, aggregate functions in DRS's AggFn enum.
//   3. Semantic validation (async) — loads metadata for every involved entity
//      and validates column / navigation property references.
//
// Errors (red) block the Apply button. Warnings (yellow) are advisory only.
// We do NOT apply state that has any errors — the builder UX would surface
// broken / phantom UI for things the user can't see or edit otherwise.

import { useMemo, useState } from 'react';
import {
  Dialog,
  DialogTrigger,
  DialogSurface,
  DialogTitle,
  DialogBody,
  DialogActions,
  DialogContent,
  Button,
  Textarea,
  Caption1,
  tokens,
  Badge,
  MessageBar,
  MessageBarBody,
  MessageBarTitle,
  Spinner,
  Link,
} from '@fluentui/react-components';
import {
  ClipboardPaste20Regular,
  ArrowDownload20Regular,
  Dismiss20Regular,
  Open16Regular,
} from '@fluentui/react-icons';
import {
  parseODataUrl,
  type ParsedRequest,
  type ParseResult,
  type ParseIssue,
} from '../engine/odataParser';
import { useScopedEntities } from '../host/useScopedEntities';
import { metadata } from '../host/metadataProvider';

export interface PasteODataDialogProps {
  /** Optional trigger element to replace the default button. Must be a single
   *  React element (DialogTrigger requires a single ReactElement child). */
  children?: React.ReactElement;
  /** Called when the user clicks Apply on a successful parse. */
  onApply: (parsed: ParsedRequest) => void;
}

export function PasteODataDialog({ children, onApply }: PasteODataDialogProps) {
  const [open, setOpen] = useState(false);
  const [text, setText] = useState('');
  const [result, setResult] = useState<ParseResult | null>(null);
  const [parsing, setParsing] = useState(false);
  const { entities } = useScopedEntities();

  const resolveEntitySet = useMemo(() => {
    return (entitySetName: string): string | undefined => {
      const hit = entities.find((e) => e.entitySetName === entitySetName);
      return hit?.logicalName;
    };
  }, [entities]);

  const onParse = async () => {
    if (!text.trim()) {
      setResult(null);
      return;
    }
    setParsing(true);
    try {
      // The third arg is the async metadata loader. The parser calls it for
      // the root entity + every $expand target so column / nav references
      // can be validated against actual metadata.
      const r = await parseODataUrl(text, resolveEntitySet, (logical) =>
        metadata.getTable(logical),
      );
      setResult(r);
    } finally {
      setParsing(false);
    }
  };

  const onApplyClick = () => {
    if (!result?.parsed) return;
    onApply(result.parsed);
    setOpen(false);
    setText('');
    setResult(null);
  };

  const onReset = () => {
    setText('');
    setResult(null);
  };

  // Apply is only enabled when the parse succeeded with zero errors AND a
  // table was resolved. Warnings don't block — they're informational.
  const canApply = !!result?.ok && !!result?.parsed?.table && result.errors.length === 0;

  return (
    <Dialog open={open} onOpenChange={(_, d) => setOpen(d.open)}>
      <DialogTrigger disableButtonEnhancement>
        {children ?? (
          <Button icon={<ClipboardPaste20Regular />} appearance="outline" size="small">
            Paste OData URL
          </Button>
        )}
      </DialogTrigger>
      <DialogSurface style={{ maxWidth: 760 }}>
        <DialogBody>
          <DialogTitle
            action={
              <DialogTrigger action="close">
                <Button appearance="subtle" icon={<Dismiss20Regular />} aria-label="close" />
              </DialogTrigger>
            }
          >
            Parse an OData URL into the builder
          </DialogTitle>
          <DialogContent>
            <Caption1
              style={{ display: 'block', marginBottom: 8, color: tokens.colorNeutralForeground2 }}
            >
              Paste a Dataverse Web API URL (absolute or relative). DRS will parse the URL, validate
              every option / function / column / nav against your environment's metadata, and only
              enable <strong>Apply</strong> if there are no errors. Warnings are advisory — they
              describe lossy but legal patterns that DRS can still represent (e.g.{' '}
              <code>not Microsoft.Dynamics.CRM.X(...)</code>
              dropping the <code>not</code>).
            </Caption1>

            <Textarea
              value={text}
              onChange={(_, d) => setText(d.value)}
              placeholder="accounts?$select=name&$filter=revenue gt 1000000&$top=10"
              style={{
                width: '100%',
                fontFamily: tokens.fontFamilyMonospace,
                fontSize: 12,
              }}
              rows={6}
              resize="vertical"
              disabled={parsing}
            />

            <div style={{ display: 'flex', gap: 8, marginTop: 12, alignItems: 'center' }}>
              <Button
                appearance="primary"
                onClick={onParse}
                disabled={!text.trim() || parsing}
                icon={parsing ? <Spinner size="tiny" /> : undefined}
              >
                {parsing ? 'Parsing & validating…' : 'Parse'}
              </Button>
              <Button
                appearance="subtle"
                onClick={onReset}
                disabled={parsing || (!text && !result)}
              >
                Clear
              </Button>
              {parsing && (
                <Caption1 style={{ color: tokens.colorNeutralForeground3 }}>
                  Loading metadata for involved entities…
                </Caption1>
              )}
            </div>

            {result && !parsing && (
              <div style={{ marginTop: 16, display: 'flex', flexDirection: 'column', gap: 8 }}>
                {/* SUCCESS — only when there were zero errors. */}
                {result.ok && result.parsed && (
                  <MessageBar layout="multiline" intent="success">
                    <MessageBarBody>
                      <MessageBarTitle>Parsed &amp; validated — ready to apply</MessageBarTitle>
                      <Summary parsed={result.parsed} />
                    </MessageBarBody>
                  </MessageBar>
                )}

                {/* ERRORS — block apply. Each issue carries an optional
                    "Learn more" link to MS docs so users can verify what
                    DRS supports without us enumerating every value. */}
                {result.errors.length > 0 && (
                  <MessageBar layout="multiline" intent="error">
                    <MessageBarBody>
                      <MessageBarTitle>
                        Won&apos;t apply — {result.errors.length} error
                        {result.errors.length === 1 ? '' : 's'} found
                      </MessageBarTitle>
                      <IssueList items={result.errors} />
                    </MessageBarBody>
                  </MessageBar>
                )}

                {/* WARNINGS — advisory only. Apply still allowed. */}
                {result.warnings.length > 0 && (
                  <MessageBar layout="multiline" intent="warning">
                    <MessageBarBody>
                      <MessageBarTitle>
                        {result.warnings.length} warning{result.warnings.length === 1 ? '' : 's'}{' '}
                        (apply still allowed)
                      </MessageBarTitle>
                      <IssueList items={result.warnings} />
                    </MessageBarBody>
                  </MessageBar>
                )}
              </div>
            )}
          </DialogContent>
          <DialogActions>
            <DialogTrigger action="close" disableButtonEnhancement>
              <Button appearance="subtle">Cancel</Button>
            </DialogTrigger>
            <Button
              appearance="primary"
              icon={<ArrowDownload20Regular />}
              onClick={onApplyClick}
              disabled={!canApply}
            >
              Apply to builder
            </Button>
          </DialogActions>
        </DialogBody>
      </DialogSurface>
    </Dialog>
  );
}

/**
 * Render a list of ParseIssue items. Each message gets a "Learn more →"
 * link beside it when the issue carries a learnMoreUrl — keeps messages
 * terse and lets the user verify spec details against MS docs directly.
 */
function IssueList({ items }: { items: ParseIssue[] }) {
  return (
    <ul style={{ margin: '4px 0 0 18px', padding: 0 }}>
      {items.map((it, i) => (
        <li key={i} style={{ marginBottom: 4 }}>
          <span>{it.message}</span>
          {it.learnMoreUrl && (
            <>
              {' '}
              <Link
                href={it.learnMoreUrl}
                target="_blank"
                rel="noopener noreferrer"
                style={{
                  display: 'inline-flex',
                  alignItems: 'center',
                  gap: 2,
                  fontSize: tokens.fontSizeBase200,
                  whiteSpace: 'nowrap',
                }}
              >
                Learn more
                <Open16Regular style={{ width: 12, height: 12 }} />
              </Link>
            </>
          )}
        </li>
      ))}
    </ul>
  );
}

function Summary({ parsed }: { parsed: ParsedRequest }) {
  return (
    <div style={{ display: 'flex', flexWrap: 'wrap', gap: 6, marginTop: 6 }}>
      <Badge appearance="tint" color="brand">
        table: <code>{parsed.table || parsed.entitySet}</code>
      </Badge>
      {parsed.select.length > 0 && <Badge appearance="tint">$select: {parsed.select.length}</Badge>}
      {parsed.filter.rules.length > 0 && (
        <Badge appearance="tint">$filter: {parsed.filter.rules.length} root nodes</Badge>
      )}
      {parsed.orderby.length > 0 && (
        <Badge appearance="tint">$orderby: {parsed.orderby.length}</Badge>
      )}
      {parsed.top != null && <Badge appearance="tint">$top: {parsed.top}</Badge>}
      {parsed.countOn && <Badge appearance="tint">$count: true</Badge>}
      {parsed.expand.length > 0 && <Badge appearance="tint">$expand: {parsed.expand.length}</Badge>}
      {parsed.apply?.enabled && <Badge appearance="tint">$apply: on</Badge>}
    </div>
  );
}
