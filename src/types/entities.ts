/**
 * CRM Entity Types
 *
 * Standard data structures for CRM entities.
 * Map your CRM's specific fields to these types in the client.
 */

// =============================================================================
// Pagination
// =============================================================================

export interface PaginationParams {
  /** Number of items to return */
  limit?: number;
  /** Cursor for pagination (CRM-specific format) */
  cursor?: string;
  /** Offset for offset-based pagination */
  offset?: number;
}

export interface PaginatedResponse<T> {
  /** Array of items */
  items: T[];
  /** Number of items in this response */
  count: number;
  /** Total count (if available) */
  total?: number;
  /** Whether more items are available */
  hasMore: boolean;
  /** Cursor for next page */
  nextCursor?: string;
}

// =============================================================================
// Search
// =============================================================================

export interface SearchParams extends PaginationParams {
  /** Search query string */
  query?: string;
  /** Filters to apply */
  filters?: SearchFilter[];
  /** Sort field */
  sortBy?: string;
  /** Sort direction */
  sortOrder?: 'asc' | 'desc';
}

export interface SearchFilter {
  /** Field to filter on */
  field: string;
  /** Filter operator */
  operator: FilterOperator;
  /** Filter value */
  value: unknown;
}

export type FilterOperator =
  | 'eq'
  | 'neq'
  | 'lt'
  | 'lte'
  | 'gt'
  | 'gte'
  | 'contains'
  | 'not_contains'
  | 'starts_with'
  | 'in'
  | 'not_in'
  | 'is_null'
  | 'is_not_null';

// =============================================================================
// Contact
// =============================================================================

export interface Contact {
  id: string;
  firstName?: string;
  lastName?: string;
  fullName?: string;
  email?: string;
  phone?: string;
  mobilePhone?: string;
  title?: string;
  department?: string;
  companyId?: string;
  companyName?: string;
  lifecycleStage?: string;
  source?: string;
  status?: string;
  address?: Address;
  createdAt?: string;
  updatedAt?: string;
  ownerId?: string;
  customFields?: Record<string, unknown>;
}

export interface ContactCreateInput {
  firstName?: string;
  lastName?: string;
  email: string;
  phone?: string;
  title?: string;
  companyId?: string;
  source?: string;
  customFields?: Record<string, unknown>;
}

export interface ContactUpdateInput {
  firstName?: string;
  lastName?: string;
  email?: string;
  phone?: string;
  title?: string;
  companyId?: string;
  customFields?: Record<string, unknown>;
}

// =============================================================================
// Company
// =============================================================================

export interface Company {
  id: string;
  name: string;
  domain?: string;
  website?: string;
  industry?: string;
  description?: string;
  numberOfEmployees?: number;
  annualRevenue?: number;
  revenueCurrency?: string;
  type?: string;
  lifecycleStage?: string;
  phone?: string;
  address?: Address;
  linkedInUrl?: string;
  createdAt?: string;
  updatedAt?: string;
  ownerId?: string;
  customFields?: Record<string, unknown>;
}

export interface CompanyCreateInput {
  name: string;
  domain?: string;
  industry?: string;
  description?: string;
  numberOfEmployees?: number;
  type?: string;
  phone?: string;
  address?: Partial<Address>;
  customFields?: Record<string, unknown>;
}

export interface CompanyUpdateInput {
  name?: string;
  domain?: string;
  industry?: string;
  description?: string;
  numberOfEmployees?: number;
  type?: string;
  phone?: string;
  address?: Partial<Address>;
  customFields?: Record<string, unknown>;
}

// =============================================================================
// Deal / Opportunity
// =============================================================================

export interface Deal {
  id: string;
  name: string;
  amount?: number;
  currency?: string;
  stage?: string;
  stageId?: string;
  pipelineId?: string;
  pipelineName?: string;
  probability?: number;
  closeDate?: string;
  status?: DealStatus;
  type?: string;
  source?: string;
  description?: string;
  companyId?: string;
  companyName?: string;
  contactIds?: string[];
  closeReason?: string;
  createdAt?: string;
  updatedAt?: string;
  ownerId?: string;
  customFields?: Record<string, unknown>;
}

export type DealStatus = 'open' | 'won' | 'lost';

export interface DealCreateInput {
  name: string;
  amount?: number;
  currency?: string;
  stageId?: string;
  pipelineId?: string;
  closeDate?: string;
  companyId?: string;
  contactIds?: string[];
  customFields?: Record<string, unknown>;
}

export interface DealUpdateInput {
  name?: string;
  amount?: number;
  currency?: string;
  stageId?: string;
  closeDate?: string;
  status?: DealStatus;
  closeReason?: string;
  customFields?: Record<string, unknown>;
}

// =============================================================================
// Pipeline
// =============================================================================

export interface Pipeline {
  id: string;
  name: string;
  stages: PipelineStage[];
  isDefault?: boolean;
}

export interface PipelineStage {
  id: string;
  name: string;
  order: number;
  probability?: number;
  isClosed?: boolean;
  isWon?: boolean;
}

// =============================================================================
// Activity
// =============================================================================

export interface Activity {
  id: string;
  type: ActivityType;
  subject: string;
  body?: string;
  status?: ActivityStatus;
  dueDate?: string;
  completedDate?: string;
  durationMinutes?: number;
  activityDate?: string;
  contactIds?: string[];
  companyId?: string;
  dealId?: string;
  createdAt?: string;
  updatedAt?: string;
  ownerId?: string;
  customFields?: Record<string, unknown>;
}

export type ActivityType = 'call' | 'email' | 'meeting' | 'task' | 'note' | 'other';
export type ActivityStatus = 'pending' | 'completed' | 'cancelled';

export interface ActivityCreateInput {
  type: ActivityType;
  subject: string;
  body?: string;
  dueDate?: string;
  contactIds?: string[];
  companyId?: string;
  dealId?: string;
  customFields?: Record<string, unknown>;
}

// =============================================================================
// Common
// =============================================================================

export interface Address {
  street?: string;
  street2?: string;
  city?: string;
  state?: string;
  postalCode?: string;
  country?: string;
}

// =============================================================================
// Lead
// =============================================================================

export interface Lead {
  id: string;
  subject?: string;
  firstName?: string;
  lastName?: string;
  fullName?: string;
  email?: string;
  phone?: string;
  mobilePhone?: string;
  companyName?: string;
  jobTitle?: string;
  website?: string;
  description?: string;
  address?: Address;
  leadSource?: number;
  leadQuality?: number;
  industryCode?: number;
  revenue?: number;
  numberOfEmployees?: number;
  status?: LeadStatus;
  stateCode?: number;
  ownerId?: string;
  parentAccountId?: string;
  parentContactId?: string;
  createdAt?: string;
  updatedAt?: string;
}

export type LeadStatus = 'open' | 'qualified' | 'disqualified';

export interface LeadCreateInput {
  subject?: string;
  firstName?: string;
  lastName: string;
  email?: string;
  phone?: string;
  companyName?: string;
  jobTitle?: string;
  website?: string;
  description?: string;
  leadSource?: number;
  leadQuality?: number;
  ownerId?: string;
  customFields?: Record<string, unknown>;
}

export interface LeadUpdateInput {
  subject?: string;
  firstName?: string;
  lastName?: string;
  email?: string;
  phone?: string;
  companyName?: string;
  jobTitle?: string;
  website?: string;
  description?: string;
  leadSource?: number;
  leadQuality?: number;
  customFields?: Record<string, unknown>;
}

export interface LeadQualifyInput {
  createAccount: boolean;
  createContact: boolean;
  createOpportunity: boolean;
  status?: number;
  opportunityCurrencyId?: string;
  opportunityCustomerId?: string;
  sourceCampaignId?: string;
}

export interface LeadQualifyResult {
  success: boolean;
  accountId?: string;
  contactId?: string;
  opportunityId?: string;
}

// =============================================================================
// Quote
// =============================================================================

export interface Quote {
  id: string;
  name: string;
  quoteNumber?: string;
  description?: string;
  effectiveFrom?: string;
  effectiveTo?: string;
  totalAmount?: number;
  totalLineItemAmount?: number;
  totalDiscountAmount?: number;
  totalTax?: number;
  freightAmount?: number;
  discountPercentage?: number;
  discountAmount?: number;
  status?: QuoteStatus;
  stateCode?: number;
  opportunityId?: string;
  customerId?: string;
  priceLevelId?: string;
  ownerId?: string;
  currencyId?: string;
  billToAddress?: Address;
  shipToAddress?: Address;
  createdAt?: string;
  updatedAt?: string;
}

export type QuoteStatus = 'draft' | 'active' | 'won' | 'closed' | 'revised';

export interface QuoteDetail {
  id: string;
  quoteId: string;
  productId?: string;
  productDescription?: string;
  quantity?: number;
  pricePerUnit?: number;
  baseAmount?: number;
  extendedAmount?: number;
  manualDiscountAmount?: number;
  tax?: number;
  uomId?: string;
  isPriceOverridden?: boolean;
  isProductOverridden?: boolean;
}

export interface QuoteCreateInput {
  name: string;
  description?: string;
  opportunityId?: string;
  customerAccountId?: string;
  customerContactId?: string;
  priceLevelId?: string;
  currencyId?: string;
  effectiveFrom?: string;
  effectiveTo?: string;
  discountPercentage?: number;
  freightAmount?: number;
  customFields?: Record<string, unknown>;
}

export interface QuoteDetailCreateInput {
  quoteId: string;
  productId?: string;
  productDescription?: string;
  quantity: number;
  pricePerUnit?: number;
  manualDiscountAmount?: number;
  tax?: number;
  uomId?: string;
  isPriceOverridden?: boolean;
}

// =============================================================================
// Sales Order
// =============================================================================

export interface SalesOrder {
  id: string;
  name: string;
  orderNumber?: string;
  description?: string;
  totalAmount?: number;
  totalLineItemAmount?: number;
  totalDiscountAmount?: number;
  totalTax?: number;
  freightAmount?: number;
  discountPercentage?: number;
  discountAmount?: number;
  dateDelivered?: string;
  requestDeliveryBy?: string;
  status?: OrderStatus;
  stateCode?: number;
  quoteId?: string;
  opportunityId?: string;
  customerId?: string;
  priceLevelId?: string;
  ownerId?: string;
  billToAddress?: Address;
  shipToAddress?: Address;
  createdAt?: string;
  updatedAt?: string;
}

export type OrderStatus = 'active' | 'submitted' | 'cancelled' | 'fulfilled' | 'invoiced';

export interface SalesOrderDetail {
  id: string;
  salesOrderId: string;
  productId?: string;
  productDescription?: string;
  quantity?: number;
  pricePerUnit?: number;
  baseAmount?: number;
  extendedAmount?: number;
  manualDiscountAmount?: number;
  tax?: number;
}

export interface SalesOrderCreateInput {
  name: string;
  description?: string;
  quoteId?: string;
  opportunityId?: string;
  customerAccountId?: string;
  priceLevelId?: string;
  requestDeliveryBy?: string;
  customFields?: Record<string, unknown>;
}

export interface SalesOrderDetailCreateInput {
  salesOrderId: string;
  productId?: string;
  productDescription?: string;
  quantity: number;
  pricePerUnit?: number;
  manualDiscountAmount?: number;
  tax?: number;
}

// =============================================================================
// Invoice
// =============================================================================

export interface Invoice {
  id: string;
  name: string;
  invoiceNumber?: string;
  description?: string;
  totalAmount?: number;
  totalLineItemAmount?: number;
  totalDiscountAmount?: number;
  totalTax?: number;
  freightAmount?: number;
  discountPercentage?: number;
  discountAmount?: number;
  dueDate?: string;
  dateDelivered?: string;
  status?: InvoiceStatus;
  stateCode?: number;
  isPriceLocked?: boolean;
  salesOrderId?: string;
  opportunityId?: string;
  customerId?: string;
  priceLevelId?: string;
  ownerId?: string;
  billToAddress?: Address;
  shipToAddress?: Address;
  createdAt?: string;
  updatedAt?: string;
}

export type InvoiceStatus = 'active' | 'closed' | 'paid' | 'cancelled';

export interface InvoiceDetail {
  id: string;
  invoiceId: string;
  productId?: string;
  productDescription?: string;
  quantity?: number;
  pricePerUnit?: number;
  baseAmount?: number;
  extendedAmount?: number;
  manualDiscountAmount?: number;
  tax?: number;
}

export interface InvoiceCreateInput {
  name: string;
  description?: string;
  salesOrderId?: string;
  customerAccountId?: string;
  priceLevelId?: string;
  dueDate?: string;
  customFields?: Record<string, unknown>;
}

// =============================================================================
// Product
// =============================================================================

export interface Product {
  id: string;
  name: string;
  productNumber?: string;
  description?: string;
  productStructure?: number; // 1=Product, 2=ProductFamily, 3=Bundle
  productTypeCode?: number; // 1=SalesInventory, 2=MiscCharges, 3=Services, 4=FlatFees
  quantityDecimal?: number;
  currentCost?: number;
  standardCost?: number;
  price?: number;
  stockWeight?: number;
  stockVolume?: number;
  quantityOnHand?: number;
  defaultUomId?: string;
  defaultUomScheduleId?: string;
  subjectId?: string;
  status?: ProductStatus;
  stateCode?: number;
  createdAt?: string;
  updatedAt?: string;
}

export type ProductStatus = 'draft' | 'active' | 'retired' | 'under_revision';

export interface ProductCreateInput {
  name: string;
  productNumber: string;
  productStructure?: number;
  productTypeCode?: number;
  quantityDecimal?: number;
  currentCost?: number;
  standardCost?: number;
  price?: number;
  defaultUomId?: string;
  defaultUomScheduleId?: string;
  description?: string;
  customFields?: Record<string, unknown>;
}

// =============================================================================
// Price List
// =============================================================================

export interface PriceLevel {
  id: string;
  name: string;
  description?: string;
  beginDate?: string;
  endDate?: string;
  freightTermsCode?: number;
  paymentMethodCode?: number;
  shippingMethodCode?: number;
  status?: string;
  stateCode?: number;
  currencyId?: string;
  createdAt?: string;
  updatedAt?: string;
}

export interface ProductPriceLevel {
  id: string;
  productId: string;
  priceLevelId: string;
  uomId?: string;
  amount?: number;
  percentage?: number;
  roundingPolicyCode?: number;
  roundingOptionCode?: number;
  roundingOptionAmount?: number;
  pricingMethodCode?: number;
}

// =============================================================================
// Competitor
// =============================================================================

export interface Competitor {
  id: string;
  name: string;
  website?: string;
  tickerSymbol?: string;
  stockExchange?: string;
  reportedRevenue?: number;
  reportingQuarter?: number;
  reportingYear?: number;
  keyProduct?: string;
  strengths?: string;
  weaknesses?: string;
  overview?: string;
  opportunities?: string;
  threats?: string;
  winPercentage?: number;
  address?: Address;
  status?: string;
  stateCode?: number;
  createdAt?: string;
  updatedAt?: string;
}

export interface CompetitorCreateInput {
  name: string;
  website?: string;
  tickerSymbol?: string;
  keyProduct?: string;
  strengths?: string;
  weaknesses?: string;
  overview?: string;
  customFields?: Record<string, unknown>;
}

// =============================================================================
// Campaign
// =============================================================================

export interface Campaign {
  id: string;
  name: string;
  codeName?: string;
  description?: string;
  message?: string;
  objective?: string;
  typeCode?: number;
  proposedStart?: string;
  proposedEnd?: string;
  actualStart?: string;
  actualEnd?: string;
  budgetedCost?: number;
  otherCost?: number;
  totalCampaignActivityActualCost?: number;
  totalActualCost?: number;
  expectedResponse?: number;
  expectedRevenue?: number;
  status?: string;
  stateCode?: number;
  priceLevelId?: string;
  ownerId?: string;
  currencyId?: string;
  createdAt?: string;
  updatedAt?: string;
}

export interface CampaignActivity {
  id: string;
  subject: string;
  description?: string;
  channelTypeCode?: number;
  budgetedCost?: number;
  actualCost?: number;
  scheduledStart?: string;
  scheduledEnd?: string;
  actualStart?: string;
  actualEnd?: string;
  status?: string;
  stateCode?: number;
  regardingObjectId?: string;
  createdAt?: string;
}

export interface CampaignResponse {
  id: string;
  subject: string;
  description?: string;
  responseCode?: number;
  receivedOn?: string;
  channelTypeCode?: number;
  promotionCodeName?: string;
  firstName?: string;
  lastName?: string;
  email?: string;
  phone?: string;
  companyName?: string;
  originatingActivityId?: string;
  regardingObjectId?: string;
  status?: string;
  stateCode?: number;
}

export interface CampaignCreateInput {
  name: string;
  codeName?: string;
  description?: string;
  message?: string;
  objective?: string;
  typeCode?: number;
  proposedStart?: string;
  proposedEnd?: string;
  budgetedCost?: number;
  priceLevelId?: string;
  currencyId?: string;
  customFields?: Record<string, unknown>;
}

// =============================================================================
// Case (Incident)
// =============================================================================

export interface Case {
  id: string;
  title: string;
  ticketNumber?: string;
  description?: string;
  caseOriginCode?: number;
  caseTypeCode?: number;
  priorityCode?: number;
  severityCode?: number;
  status?: CaseStatus;
  stateCode?: number;
  escalatedOn?: string;
  isEscalated?: boolean;
  followUpBy?: string;
  customerId?: string;
  primaryContactId?: string;
  productId?: string;
  subjectId?: string;
  ownerId?: string;
  entitlementId?: string;
  contractId?: string;
  contractDetailId?: string;
  resolveBy?: string;
  responseBy?: string;
  createdAt?: string;
  updatedAt?: string;
}

export type CaseStatus = 'active' | 'resolved' | 'cancelled';

export interface CaseResolution {
  id: string;
  subject: string;
  description?: string;
  timeSpent?: number;
  totalTime?: number;
  incidentId?: string;
  createdAt?: string;
}

export interface CaseCreateInput {
  title: string;
  description?: string;
  caseOriginCode?: number;
  caseTypeCode?: number;
  priorityCode?: number;
  severityCode?: number;
  customerAccountId?: string;
  customerContactId?: string;
  primaryContactId?: string;
  productId?: string;
  subjectId?: string;
  customFields?: Record<string, unknown>;
}

// =============================================================================
// Goal
// =============================================================================

export interface Goal {
  id: string;
  title: string;
  goalOwnerId?: string;
  metricId?: string;
  targetMoney?: number;
  targetDecimal?: number;
  targetInteger?: number;
  actualMoney?: number;
  actualDecimal?: number;
  actualInteger?: number;
  inProgressMoney?: number;
  inProgressDecimal?: number;
  inProgressInteger?: number;
  percentage?: number;
  fiscalPeriod?: number;
  fiscalYear?: number;
  goalStartDate?: string;
  goalEndDate?: string;
  considerOnlyGoalOwnersRecords?: boolean;
  parentGoalId?: string;
  rollupQueryActualMoneyId?: string;
  status?: string;
  stateCode?: number;
  lastRolledUpDate?: string;
  createdAt?: string;
  updatedAt?: string;
}

export interface GoalMetric {
  id: string;
  name: string;
  description?: string;
  amountDataType?: number;
  isAmount?: boolean;
  isStretchTracked?: boolean;
  status?: string;
  stateCode?: number;
}

export interface GoalCreateInput {
  title: string;
  goalOwnerId: string;
  metricId: string;
  targetMoney?: number;
  targetDecimal?: number;
  targetInteger?: number;
  goalStartDate?: string;
  goalEndDate?: string;
  fiscalPeriod?: number;
  fiscalYear?: number;
  considerOnlyGoalOwnersRecords?: boolean;
  parentGoalId?: string;
  customFields?: Record<string, unknown>;
}

// =============================================================================
// Note (Annotation)
// =============================================================================

export interface Note {
  id: string;
  subject?: string;
  noteText?: string;
  objectId?: string;
  objectTypeCode?: string;
  isDocument?: boolean;
  fileName?: string;
  mimeType?: string;
  fileSize?: number;
  documentBody?: string;
  ownerId?: string;
  createdById?: string;
  createdAt?: string;
  updatedAt?: string;
}

export interface NoteCreateInput {
  subject?: string;
  noteText?: string;
  // Legacy format
  objectType?: 'account' | 'contact' | 'opportunity' | 'lead' | 'incident' | 'quote' | 'salesorder' | 'invoice';
  objectId?: string;
  // New format
  regardingEntityType?: string;
  regardingId?: string;
  isDocument?: boolean;
  fileName?: string;
  mimeType?: string;
  documentBody?: string;
  customFields?: Record<string, unknown>;
}

// =============================================================================
// User and Team
// =============================================================================

export interface SystemUser {
  id: string;
  fullName?: string;
  firstName?: string;
  lastName?: string;
  domainName?: string;
  internalEmailAddress?: string;
  title?: string;
  jobTitle?: string;
  mobilePhone?: string;
  phone?: string;
  businessUnitId?: string;
  territoryId?: string;
  positionId?: string;
  queueId?: string;
  isDisabled?: boolean;
  accessMode?: number;
  userLicenseType?: number;
  setupUser?: boolean;
  createdAt?: string;
  updatedAt?: string;
}

export interface Team {
  id: string;
  name: string;
  description?: string;
  teamType?: number; // 0=Owner, 1=Access, 2=AAD Security Group, 3=AAD Office Group
  businessUnitId?: string;
  administratorId?: string;
  queueId?: string;
  isDefault?: boolean;
  createdAt?: string;
}

export interface BusinessUnit {
  id: string;
  name: string;
  parentBusinessUnitId?: string;
  divisionName?: string;
  emailAddress?: string;
  website?: string;
  isDisabled?: boolean;
  createdAt?: string;
}

// =============================================================================
// Metadata
// =============================================================================

export interface EntityMetadata {
  logicalName: string;
  displayName: string;
  description?: string;
  primaryIdAttribute: string;
  primaryNameAttribute: string;
  entitySetName: string;
  isCustomEntity?: boolean;
  isValidForAdvancedFind?: boolean;
  canCreateAttributes?: boolean;
  attributes?: AttributeMetadata[];
}

export interface AttributeMetadata {
  logicalName: string;
  displayName: string;
  description?: string;
  attributeType: string;
  requiredLevel: string;
  maxLength?: number;
  minValue?: number;
  maxValue?: number;
  options?: OptionMetadata[];
}

export interface OptionMetadata {
  value: number;
  label: string;
}

// =============================================================================
// Response Format
// =============================================================================

export type ResponseFormat = 'json' | 'markdown';
