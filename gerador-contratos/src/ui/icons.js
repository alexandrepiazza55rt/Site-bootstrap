// Ícones (Lucide, licença ISC) — apenas os usados, para manter o app leve.
import { h } from 'preact';
import {
  FileText, Files, Plus, Search, House, LayoutDashboard, Folder, FolderOpen, Settings, Users, User, Building2, BookOpen,
  Library, Download, Printer, FileDown, Undo2, Redo2, ZoomIn, ZoomOut, Eye, EyeOff, Highlighter, ChevronLeft,
  ChevronRight, ChevronDown, ChevronUp, ArrowUp, ArrowDown, Copy, Trash2, Pencil, X, Check, CircleCheck,
  TriangleAlert, Info, WandSparkles, Sparkles, Calendar, CalendarPlus, Calculator, Hash, Coins, Type, AlignLeft,
  AlignCenter, AlignRight, AlignJustify, Bold, Italic, Underline, List, ListOrdered, Heading, Pilcrow, MapPin, Phone,
  Mail, Tag, Truck, Upload, Save, Ellipsis, EllipsisVertical, RotateCcw, History, GitCompare, Clock, Braces,
  TextCursorInput, FilePlus, Star, Moon, Sun, Monitor, HardDrive, Database, ShieldCheck, GripVertical, ExternalLink,
  Lock, Layers, ListTree, BookmarkPlus, Package, IdCard, Menu, PanelLeft, PanelRight, Maximize2, Keyboard, Lightbulb,
  Rocket, PartyPopper, CalendarClock, CircleAlert, FileCheck2, Signature, Scale, Receipt, Briefcase, CopyPlus,
  ArrowRight, ArrowLeft, CircleHelp, Share2, SquarePen, MousePointerClick, Replace, BadgeCheck, SearchCheck, Filter,
  ListPlus, SeparatorHorizontal, Section,
} from 'lucide';

const MAP = {
  'file-text': FileText, files: Files, plus: Plus, search: Search, home: House, dashboard: LayoutDashboard,
  folder: Folder, 'folder-open': FolderOpen, settings: Settings, users: Users, user: User, building: Building2,
  book: BookOpen, library: Library, download: Download, printer: Printer, 'file-down': FileDown, undo: Undo2,
  redo: Redo2, 'zoom-in': ZoomIn, 'zoom-out': ZoomOut, eye: Eye, 'eye-off': EyeOff, highlighter: Highlighter,
  'chevron-left': ChevronLeft, 'chevron-right': ChevronRight, 'chevron-down': ChevronDown, 'chevron-up': ChevronUp,
  'arrow-up': ArrowUp, 'arrow-down': ArrowDown, copy: Copy, trash: Trash2, pencil: Pencil, x: X, check: Check,
  'check-circle': CircleCheck, alert: TriangleAlert, info: Info, wand: WandSparkles, sparkles: Sparkles,
  calendar: Calendar, 'calendar-plus': CalendarPlus, calculator: Calculator, hash: Hash, coins: Coins, type: Type,
  'align-left': AlignLeft, 'align-center': AlignCenter, 'align-right': AlignRight, 'align-justify': AlignJustify,
  bold: Bold, italic: Italic, underline: Underline, list: List, 'list-ordered': ListOrdered, 'list-letters': ListOrdered,
  heading: Heading, pilcrow: Pilcrow, 'map-pin': MapPin, phone: Phone, mail: Mail, tag: Tag, truck: Truck,
  upload: Upload, save: Save, more: Ellipsis, 'more-v': EllipsisVertical, reset: RotateCcw, history: History,
  compare: GitCompare, clock: Clock, braces: Braces, 'text-input': TextCursorInput, 'file-plus': FilePlus, star: Star,
  moon: Moon, sun: Sun, monitor: Monitor, drive: HardDrive, database: Database, shield: ShieldCheck,
  grip: GripVertical, external: ExternalLink, lock: Lock, layers: Layers, tree: ListTree, 'bookmark-plus': BookmarkPlus,
  package: Package, id: IdCard, menu: Menu, 'panel-left': PanelLeft, 'panel-right': PanelRight, maximize: Maximize2,
  keyboard: Keyboard, lightbulb: Lightbulb, rocket: Rocket, party: PartyPopper, 'calendar-clock': CalendarClock,
  'alert-circle': CircleAlert, 'file-check': FileCheck2, signature: Signature, scale: Scale, receipt: Receipt,
  briefcase: Briefcase, 'copy-plus': CopyPlus, 'arrow-right': ArrowRight, 'arrow-left': ArrowLeft, help: CircleHelp,
  share: Share2, edit: SquarePen, click: MousePointerClick, replace: Replace, badge: BadgeCheck,
  'search-check': SearchCheck, filter: Filter, 'list-plus': ListPlus, 'page-break': SeparatorHorizontal, section: Section,
};

export function Icon({ name, size = 18, class: cls = '', title, ...rest }) {
  const node = MAP[name] || MAP['file-text'];
  const [, attrs, children] = node;
  return h('svg', {
    ...attrs, width: size, height: size, class: 'icon ' + cls, 'aria-hidden': title ? undefined : 'true',
    role: title ? 'img' : undefined, ...rest,
  }, [title ? h('title', null, title) : null, ...children.map(([tag, a], i) => h(tag, { key: i, ...a }))]);
}
