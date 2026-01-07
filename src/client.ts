/**
 * Dynamics 365 Web API Client
 *
 * Implements the CRM client interface for Microsoft Dynamics 365.
 * Reference: https://learn.microsoft.com/en-us/power-apps/developer/data-platform/webapi/overview
 *
 * MULTI-TENANT: This client receives credentials per-request via TenantCredentials,
 * allowing a single server to serve multiple tenants with different credentials.
 */

import type {
  Activity,
  ActivityCreateInput,
  BusinessUnit,
  Campaign,
  CampaignCreateInput,
  Case,
  CaseCreateInput,
  CaseStatus,
  Company,
  CompanyCreateInput,
  CompanyUpdateInput,
  Competitor,
  CompetitorCreateInput,
  Contact,
  ContactCreateInput,
  ContactUpdateInput,
  Deal,
  DealCreateInput,
  DealUpdateInput,
  Goal,
  GoalCreateInput,
  GoalMetric,
  Invoice,
  InvoiceCreateInput,
  InvoiceStatus,
  Lead,
  LeadCreateInput,
  LeadUpdateInput,
  LeadQualifyInput,
  LeadQualifyResult,
  Note,
  NoteCreateInput,
  OrderStatus,
  PaginatedResponse,
  PaginationParams,
  Pipeline,
  PipelineStage,
  PriceLevel,
  Product,
  ProductCreateInput,
  ProductStatus,
  Quote,
  QuoteCreateInput,
  QuoteDetail,
  QuoteDetailCreateInput,
  QuoteStatus,
  SalesOrder,
  SalesOrderCreateInput,
  SalesOrderDetail,
  SalesOrderDetailCreateInput,
  SearchParams,
  SystemUser,
  Team,
} from './types/entities.js';
import type { TenantCredentials } from './types/env.js';
import { AuthenticationError, CrmApiError, RateLimitError } from './utils/errors.js';

// =============================================================================
// Dynamics 365 Response Types
// =============================================================================

interface DynamicsODataResponse<T> {
  '@odata.context'?: string;
  '@odata.count'?: number;
  '@odata.nextLink'?: string;
  value: T[];
}

interface DynamicsContact {
  contactid: string;
  firstname?: string;
  lastname?: string;
  fullname?: string;
  emailaddress1?: string;
  emailaddress2?: string;
  emailaddress3?: string;
  telephone1?: string;
  telephone2?: string;
  mobilephone?: string;
  jobtitle?: string;
  department?: string;
  _parentcustomerid_value?: string;
  createdon?: string;
  modifiedon?: string;
  statecode?: number;
  _ownerid_value?: string;
}

interface DynamicsAccount {
  accountid: string;
  name: string;
  websiteurl?: string;
  industrycode?: number;
  description?: string;
  numberofemployees?: number;
  revenue?: number;
  telephone1?: string;
  address1_line1?: string;
  address1_city?: string;
  address1_stateorprovince?: string;
  address1_country?: string;
  address1_postalcode?: string;
  createdon?: string;
  modifiedon?: string;
  statecode?: number;
  _ownerid_value?: string;
}

interface DynamicsOpportunity {
  opportunityid: string;
  name: string;
  estimatedvalue?: number;
  _transactioncurrencyid_value?: string;
  stepname?: string;
  estimatedclosedate?: string;
  closeprobability?: number;
  _parentaccountid_value?: string;
  _parentcontactid_value?: string;
  description?: string;
  statecode?: number;
  statuscode?: number;
  createdon?: string;
  modifiedon?: string;
  _ownerid_value?: string;
}

interface DynamicsWorkflow {
  workflowid: string;
  name: string;
  category?: number;
  statecode?: number;
}

interface DynamicsProcessStage {
  processstageid: string;
  stagename: string;
  _processid_value: string;
  stagecategory?: number;
}

interface DynamicsActivityPointer {
  activityid: string;
  subject?: string;
  description?: string;
  activitytypecode?: string;
  scheduledstart?: string;
  scheduledend?: string;
  actualend?: string;
  actualdurationminutes?: number;
  statecode?: number;
  statuscode?: number;
  createdon?: string;
  modifiedon?: string;
  _regardingobjectid_value?: string;
  _ownerid_value?: string;
}

interface DynamicsLead {
  leadid: string;
  subject?: string;
  firstname?: string;
  lastname?: string;
  fullname?: string;
  emailaddress1?: string;
  telephone1?: string;
  mobilephone?: string;
  companyname?: string;
  jobtitle?: string;
  websiteurl?: string;
  address1_line1?: string;
  address1_city?: string;
  address1_stateorprovince?: string;
  address1_postalcode?: string;
  address1_country?: string;
  description?: string;
  leadsourcecode?: number;
  leadqualitycode?: number;
  industrycode?: number;
  revenue?: number;
  numberofemployees?: number;
  statuscode?: number;
  statecode?: number;
  _ownerid_value?: string;
  _parentaccountid_value?: string;
  _parentcontactid_value?: string;
  createdon?: string;
  modifiedon?: string;
}

interface DynamicsQuote {
  quoteid: string;
  name: string;
  quotenumber?: string;
  description?: string;
  effectivefrom?: string;
  effectiveto?: string;
  totalamount?: number;
  totallineitemamount?: number;
  totaldiscountamount?: number;
  totaltax?: number;
  freightamount?: number;
  discountpercentage?: number;
  discountamount?: number;
  statuscode?: number;
  statecode?: number;
  _opportunityid_value?: string;
  _customerid_value?: string;
  _pricelevelid_value?: string;
  _ownerid_value?: string;
  _transactioncurrencyid_value?: string;
  createdon?: string;
  modifiedon?: string;
}

interface DynamicsQuoteDetail {
  quotedetailid: string;
  _quoteid_value: string;
  _productid_value?: string;
  productdescription?: string;
  quantity?: number;
  priceperunit?: number;
  baseamount?: number;
  extendedamount?: number;
  manualdiscountamount?: number;
  tax?: number;
  _uomid_value?: string;
  ispriceoverridden?: boolean;
  isproductoverridden?: boolean;
}

interface DynamicsSalesOrder {
  salesorderid: string;
  name: string;
  ordernumber?: string;
  description?: string;
  totalamount?: number;
  totallineitemamount?: number;
  totaldiscountamount?: number;
  totaltax?: number;
  freightamount?: number;
  discountpercentage?: number;
  discountamount?: number;
  datedelivered?: string;
  requestdeliveryby?: string;
  statuscode?: number;
  statecode?: number;
  _quoteid_value?: string;
  _opportunityid_value?: string;
  _customerid_value?: string;
  _pricelevelid_value?: string;
  _ownerid_value?: string;
  createdon?: string;
  modifiedon?: string;
}

interface DynamicsInvoice {
  invoiceid: string;
  name: string;
  invoicenumber?: string;
  description?: string;
  totalamount?: number;
  totallineitemamount?: number;
  totaldiscountamount?: number;
  totaltax?: number;
  freightamount?: number;
  discountpercentage?: number;
  discountamount?: number;
  duedate?: string;
  datedelivered?: string;
  statuscode?: number;
  statecode?: number;
  ispricelocked?: boolean;
  _salesorderid_value?: string;
  _opportunityid_value?: string;
  _customerid_value?: string;
  _pricelevelid_value?: string;
  _ownerid_value?: string;
  createdon?: string;
  modifiedon?: string;
}

interface DynamicsProduct {
  productid: string;
  name: string;
  productnumber?: string;
  description?: string;
  productstructure?: number;
  producttypecode?: number;
  quantitydecimal?: number;
  currentcost?: number;
  standardcost?: number;
  price?: number;
  stockweight?: number;
  stockvolume?: number;
  quantityonhand?: number;
  _defaultuomid_value?: string;
  _defaultuomscheduleid_value?: string;
  _subjectid_value?: string;
  statuscode?: number;
  statecode?: number;
  createdon?: string;
  modifiedon?: string;
}

interface DynamicsPriceLevel {
  pricelevelid: string;
  name: string;
  description?: string;
  begindate?: string;
  enddate?: string;
  freighttermscode?: number;
  paymentmethodcode?: number;
  shippingmethodcode?: number;
  statuscode?: number;
  statecode?: number;
  _transactioncurrencyid_value?: string;
  createdon?: string;
  modifiedon?: string;
}

interface DynamicsCompetitor {
  competitorid: string;
  name: string;
  websiteurl?: string;
  tickersymbol?: string;
  stockexchange?: string;
  reportedrevenue?: number;
  reportingquarter?: number;
  reportingyear?: number;
  keyproduct?: string;
  strengths?: string;
  weaknesses?: string;
  overview?: string;
  opportunities?: string;
  threats?: string;
  winpercentage?: number;
  address1_line1?: string;
  address1_city?: string;
  address1_stateorprovince?: string;
  address1_postalcode?: string;
  address1_country?: string;
  statuscode?: number;
  statecode?: number;
  createdon?: string;
  modifiedon?: string;
}

interface DynamicsCampaign {
  campaignid: string;
  name: string;
  codename?: string;
  description?: string;
  message?: string;
  objective?: string;
  typecode?: number;
  proposedstart?: string;
  proposedend?: string;
  actualstart?: string;
  actualend?: string;
  budgetedcost?: number;
  othercost?: number;
  totalcampaignactivityactualcost?: number;
  totalactualcost?: number;
  expectedresponse?: number;
  expectedrevenue?: number;
  statuscode?: number;
  statecode?: number;
  _pricelevelid_value?: string;
  _ownerid_value?: string;
  _transactioncurrencyid_value?: string;
  createdon?: string;
  modifiedon?: string;
}

interface DynamicsIncident {
  incidentid: string;
  title: string;
  ticketnumber?: string;
  description?: string;
  caseorigincode?: number;
  casetypecode?: number;
  prioritycode?: number;
  severitycode?: number;
  statuscode?: number;
  statecode?: number;
  escalatedon?: string;
  isescalated?: boolean;
  followupby?: string;
  _customerid_value?: string;
  _primarycontactid_value?: string;
  _productid_value?: string;
  _subjectid_value?: string;
  _ownerid_value?: string;
  _entitlementid_value?: string;
  _contractid_value?: string;
  _contractdetailid_value?: string;
  resolveby?: string;
  responseby?: string;
  createdon?: string;
  modifiedon?: string;
}

interface DynamicsGoal {
  goalid: string;
  title: string;
  _goalownerid_value?: string;
  _metricid_value?: string;
  targetmoney?: number;
  targetdecimal?: number;
  targetinteger?: number;
  actualmoney?: number;
  actualdecimal?: number;
  actualinteger?: number;
  inprogressmoney?: number;
  inprogressdecimal?: number;
  inprogressinteger?: number;
  percentage?: number;
  fiscalperiod?: number;
  fiscalyear?: number;
  goalstartdate?: string;
  goalenddate?: string;
  consideronlygoalownersrecords?: boolean;
  _parentgoalid_value?: string;
  statuscode?: number;
  statecode?: number;
  lastrolledupdate?: string;
  createdon?: string;
  modifiedon?: string;
}

interface DynamicsAnnotation {
  annotationid: string;
  subject?: string;
  notetext?: string;
  _objectid_value?: string;
  objecttypecode?: string;
  isdocument?: boolean;
  filename?: string;
  mimetype?: string;
  filesize?: number;
  documentbody?: string;
  _ownerid_value?: string;
  _createdby_value?: string;
  createdon?: string;
  modifiedon?: string;
}

interface DynamicsSystemUser {
  systemuserid: string;
  fullname?: string;
  firstname?: string;
  lastname?: string;
  domainname?: string;
  internalemailaddress?: string;
  title?: string;
  jobtitle?: string;
  mobilephone?: string;
  address1_telephone1?: string;
  _businessunitid_value?: string;
  _territoryid_value?: string;
  _positionid_value?: string;
  _queueid_value?: string;
  isdisabled?: boolean;
  accessmode?: number;
  userlicensetype?: number;
  setupuser?: boolean;
  createdon?: string;
  modifiedon?: string;
}

interface DynamicsTeam {
  teamid: string;
  name: string;
  description?: string;
  teamtype?: number;
  _businessunitid_value?: string;
  _administratorid_value?: string;
  _queueid_value?: string;
  isdefault?: boolean;
  createdon?: string;
}

interface DynamicsBusinessUnit {
  businessunitid: string;
  name: string;
  _parentbusinessunitid_value?: string;
  divisionname?: string;
  emailaddress?: string;
  websiteurl?: string;
  isdisabled?: boolean;
  createdon?: string;
}

interface DynamicsWhoAmI {
  UserId: string;
  BusinessUnitId: string;
  OrganizationId: string;
}

// =============================================================================
// CRM Client Interface
// =============================================================================

export interface CrmClient {
  // Connection
  testConnection(): Promise<{ connected: boolean; message: string }>;

  // Contacts
  listContacts(params?: PaginationParams): Promise<PaginatedResponse<Contact>>;
  getContact(id: string): Promise<Contact>;
  createContact(input: ContactCreateInput): Promise<Contact>;
  updateContact(id: string, input: ContactUpdateInput): Promise<Contact>;
  deleteContact(id: string): Promise<void>;
  searchContacts(params: SearchParams): Promise<PaginatedResponse<Contact>>;

  // Companies
  listCompanies(params?: PaginationParams): Promise<PaginatedResponse<Company>>;
  getCompany(id: string): Promise<Company>;
  createCompany(input: CompanyCreateInput): Promise<Company>;
  updateCompany(id: string, input: CompanyUpdateInput): Promise<Company>;

  // Deals
  listDeals(params?: PaginationParams): Promise<PaginatedResponse<Deal>>;
  getDeal(id: string): Promise<Deal>;
  createDeal(input: DealCreateInput): Promise<Deal>;
  updateDeal(id: string, input: DealUpdateInput): Promise<Deal>;
  moveDealStage(id: string, stageId: string): Promise<Deal>;
  listPipelines(): Promise<Pipeline[]>;

  // Activities
  listActivities(
    params?: PaginationParams & { recordId?: string }
  ): Promise<PaginatedResponse<Activity>>;
  createActivity(input: ActivityCreateInput): Promise<Activity>;
  getActivity(id: string, activityType?: string): Promise<Activity>;
  updateActivity(id: string, activityType: string, input: { subject?: string; description?: string; scheduledStart?: string; scheduledEnd?: string; priorityCode?: number }): Promise<Activity>;
  deleteActivity(id: string, activityType: string): Promise<void>;
  completeActivity(id: string, activityType: string): Promise<void>;
  cancelActivity(id: string, activityType: string): Promise<void>;
  logCall(
    contactId: string,
    subject: string,
    notes?: string,
    durationMinutes?: number
  ): Promise<Activity>;
  logEmail(
    contactId: string,
    subject: string,
    body: string,
    direction: 'sent' | 'received'
  ): Promise<Activity>;
  createAppointment(input: {
    subject: string;
    description?: string;
    location?: string;
    scheduledStart: string;
    scheduledEnd: string;
    regardingId?: string;
    regardingType?: string;
    requiredAttendees?: string[];
    optionalAttendees?: string[];
    isAllDayEvent?: boolean;
  }): Promise<Activity>;
  createLetter(input: {
    subject: string;
    description?: string;
    regardingId?: string;
    regardingType?: string;
    address?: string;
  }): Promise<Activity>;
  createFax(input: {
    subject: string;
    description?: string;
    regardingId?: string;
    regardingType?: string;
    faxNumber?: string;
  }): Promise<Activity>;
  listActivitiesByType(activityType: string, params?: PaginationParams & { regardingId?: string }): Promise<PaginatedResponse<Activity>>;
  getActivityParties(activityId: string): Promise<unknown[]>;

  // Leads
  listLeads(params?: PaginationParams): Promise<PaginatedResponse<Lead>>;
  getLead(id: string): Promise<Lead>;
  createLead(input: LeadCreateInput): Promise<Lead>;
  updateLead(id: string, input: LeadUpdateInput): Promise<Lead>;
  deleteLead(id: string): Promise<void>;
  searchLeads(params: SearchParams): Promise<PaginatedResponse<Lead>>;
  qualifyLead(id: string, input: LeadQualifyInput): Promise<LeadQualifyResult>;

  // Quotes
  listQuotes(params?: PaginationParams): Promise<PaginatedResponse<Quote>>;
  getQuote(id: string): Promise<Quote>;
  createQuote(input: QuoteCreateInput): Promise<Quote>;
  updateQuote(id: string, input: Partial<QuoteCreateInput>): Promise<Quote>;
  deleteQuote(id: string): Promise<void>;
  listQuoteDetails(quoteId: string): Promise<QuoteDetail[]>;
  addQuoteDetail(input: QuoteDetailCreateInput): Promise<QuoteDetail>;
  activateQuote(id: string): Promise<void>;
  closeQuote(id: string, status: 'won' | 'lost' | 'cancelled'): Promise<void>;
  convertQuoteToOrder(quoteId: string): Promise<{ salesOrderId: string }>;

  // Sales Orders
  listOrders(params?: PaginationParams): Promise<PaginatedResponse<SalesOrder>>;
  getOrder(id: string): Promise<SalesOrder>;
  createOrder(input: SalesOrderCreateInput): Promise<SalesOrder>;
  updateOrder(id: string, input: Partial<SalesOrderCreateInput>): Promise<SalesOrder>;
  deleteOrder(id: string): Promise<void>;
  listOrderDetails(orderId: string): Promise<SalesOrderDetail[]>;
  addOrderDetail(input: SalesOrderDetailCreateInput): Promise<SalesOrderDetail>;
  fulfillOrder(id: string): Promise<void>;
  cancelOrder(id: string): Promise<void>;
  convertOrderToInvoice(orderId: string): Promise<{ invoiceId: string }>;

  // Invoices
  listInvoices(params?: PaginationParams): Promise<PaginatedResponse<Invoice>>;
  getInvoice(id: string): Promise<Invoice>;
  createInvoice(input: InvoiceCreateInput): Promise<Invoice>;
  updateInvoice(id: string, input: Partial<InvoiceCreateInput>): Promise<Invoice>;
  deleteInvoice(id: string): Promise<void>;
  lockInvoicePricing(id: string): Promise<void>;
  cancelInvoice(id: string): Promise<void>;

  // Products
  listProducts(params?: PaginationParams): Promise<PaginatedResponse<Product>>;
  getProduct(id: string): Promise<Product>;
  createProduct(input: ProductCreateInput): Promise<Product>;
  updateProduct(id: string, input: Partial<ProductCreateInput>): Promise<Product>;
  deleteProduct(id: string): Promise<void>;
  publishProduct(id: string): Promise<void>;

  // Price Lists
  listPriceLists(params?: PaginationParams): Promise<PaginatedResponse<PriceLevel>>;
  getPriceList(id: string): Promise<PriceLevel>;
  createPriceList(input: { name: string; description?: string; currencyId?: string }): Promise<PriceLevel>;

  // Competitors
  listCompetitors(params?: PaginationParams): Promise<PaginatedResponse<Competitor>>;
  getCompetitor(id: string): Promise<Competitor>;
  createCompetitor(input: CompetitorCreateInput): Promise<Competitor>;
  updateCompetitor(id: string, input: Partial<CompetitorCreateInput>): Promise<Competitor>;
  deleteCompetitor(id: string): Promise<void>;
  associateCompetitorToOpportunity(competitorId: string, opportunityId: string): Promise<void>;
  disassociateCompetitorFromOpportunity(competitorId: string, opportunityId: string): Promise<void>;
  listOpportunityCompetitors(opportunityId: string): Promise<Competitor[]>;

  // Campaigns
  listCampaigns(params?: PaginationParams): Promise<PaginatedResponse<Campaign>>;
  getCampaign(id: string): Promise<Campaign>;
  createCampaign(input: CampaignCreateInput): Promise<Campaign>;
  updateCampaign(id: string, input: Partial<CampaignCreateInput>): Promise<Campaign>;
  deleteCampaign(id: string): Promise<void>;
  listCampaignActivities(campaignId: string): Promise<unknown[]>;
  createCampaignActivity(input: {
    campaignId: string;
    subject: string;
    description?: string;
    channelTypeCode?: number;
    typeCode?: number;
    scheduledStart?: string;
    scheduledEnd?: string;
    budgetedCost?: number;
  }): Promise<unknown>;
  listCampaignResponses(campaignId: string): Promise<unknown[]>;
  createCampaignResponse(input: {
    campaignId: string;
    subject: string;
    description?: string;
    channelTypeCode?: number;
    responseCode?: number;
    receivedOn?: string;
    firstName?: string;
    lastName?: string;
    emailAddress?: string;
    telephone?: string;
    companyName?: string;
  }): Promise<unknown>;
  addCampaignMember(campaignId: string, entityType: string, entityId: string): Promise<void>;
  removeCampaignMember(campaignId: string, entityType: string, entityId: string): Promise<void>;

  // Cases
  listCases(params?: PaginationParams): Promise<PaginatedResponse<Case>>;
  getCase(id: string): Promise<Case>;
  createCase(input: CaseCreateInput): Promise<Case>;
  updateCase(id: string, input: Partial<CaseCreateInput>): Promise<Case>;
  deleteCase(id: string): Promise<void>;
  resolveCase(id: string, resolution: { subject: string; description?: string; billableTime?: number; resolution?: string }): Promise<void>;
  cancelCase(id: string): Promise<void>;
  reactivateCase(id: string): Promise<void>;

  // Goals
  listGoals(params?: PaginationParams): Promise<PaginatedResponse<Goal>>;
  getGoal(id: string): Promise<Goal>;
  createGoal(input: GoalCreateInput): Promise<Goal>;
  updateGoal(id: string, input: Partial<GoalCreateInput>): Promise<Goal>;
  deleteGoal(id: string): Promise<void>;
  recalculateGoal(id: string): Promise<void>;
  listGoalMetrics(params?: PaginationParams): Promise<PaginatedResponse<GoalMetric>>;
  getGoalMetric(id: string): Promise<GoalMetric>;

  // Notes
  listNotes(params?: PaginationParams & { regardingEntityType?: string; regardingId?: string }): Promise<PaginatedResponse<Note>>;
  getNote(id: string): Promise<Note>;
  createNote(input: NoteCreateInput): Promise<Note>;
  updateNote(id: string, input: { subject?: string; noteText?: string; fileName?: string; mimeType?: string; documentBody?: string }): Promise<Note>;
  deleteNote(id: string): Promise<void>;
  getNoteAttachment(id: string): Promise<{ fileName: string; mimeType: string; documentBody: string; fileSize: number }>;
  listEntityNotes(entityType: string, entityId: string, limit?: number): Promise<Note[]>;
  addAttachmentToNote(noteId: string, input: { fileName: string; mimeType: string; documentBody: string }): Promise<void>;
  removeAttachmentFromNote(noteId: string): Promise<void>;
  searchNotes(query: string, limit?: number): Promise<Note[]>;

  // Users
  listUsers(params?: PaginationParams & { activeOnly?: boolean }): Promise<PaginatedResponse<SystemUser>>;
  getUser(id: string): Promise<SystemUser>;
  getCurrentUser(): Promise<SystemUser>;
  searchUsers(query: string, limit?: number): Promise<SystemUser[]>;

  // Teams
  listTeams(params?: PaginationParams): Promise<PaginatedResponse<Team>>;
  getTeam(id: string): Promise<Team>;
  createTeam(input: { name: string; description?: string; businessUnitId: string; teamType?: number; administratorId?: string }): Promise<Team>;
  updateTeam(id: string, input: { name?: string; description?: string; administratorId?: string }): Promise<Team>;
  deleteTeam(id: string): Promise<void>;
  addTeamMember(teamId: string, userId: string): Promise<void>;
  removeTeamMember(teamId: string, userId: string): Promise<void>;
  listTeamMembers(teamId: string): Promise<SystemUser[]>;

  // Business Units
  listBusinessUnits(params?: PaginationParams): Promise<PaginatedResponse<BusinessUnit>>;
  getBusinessUnit(id: string): Promise<BusinessUnit>;

  // Dynamics-specific
  executeFetchXml(entity: string, fetchXml: string): Promise<unknown[]>;

  // Batch operations
  batchCreate(entityType: string, records: Record<string, unknown>[], continueOnError?: boolean): Promise<{
    succeeded: number;
    failed: number;
    errors: { index: number; message: string }[];
    createdIds: string[];
  }>;
  batchUpdate(entityType: string, updates: { id: string; data: Record<string, unknown> }[], continueOnError?: boolean): Promise<{
    succeeded: number;
    failed: number;
    errors: { index: number; message: string }[];
  }>;
  batchDelete(entityType: string, ids: string[], continueOnError?: boolean): Promise<{
    succeeded: number;
    failed: number;
    errors: { index: number; message: string }[];
  }>;
  batchUpsert(entityType: string, records: { alternateKey?: Record<string, string>; id?: string; data: Record<string, unknown> }[], continueOnError?: boolean): Promise<{
    succeeded: number;
    failed: number;
    created: number;
    updated: number;
    errors: { index: number; message: string }[];
  }>;

  // OData Query
  executeQuery(params: {
    entityType: string;
    select?: string;
    filter?: string;
    orderby?: string;
    expand?: string;
    top?: number;
    skip?: number;
    count?: boolean;
  }): Promise<{ records: unknown[]; totalCount?: number; hasMore: boolean }>;
  executeAggregate(params: {
    entityType: string;
    aggregateFunction: 'count' | 'sum' | 'avg' | 'min' | 'max';
    field?: string;
    filter?: string;
    groupBy?: string;
  }): Promise<unknown>;
  getRecordCount(entityType: string, filter?: string): Promise<number>;
  executeSavedQuery(savedQueryId: string, entityType: string, top?: number): Promise<unknown[]>;
  executeUserQuery(userQueryId: string, entityType: string, top?: number): Promise<unknown[]>;

  // Metadata
  listEntityMetadata(params?: { filter?: string; includeCustom?: boolean }): Promise<{
    logicalName: string;
    displayName: string;
    description?: string;
    isCustomEntity: boolean;
    primaryIdAttribute: string;
    primaryNameAttribute: string;
  }[]>;
  getEntityMetadata(entityLogicalName: string): Promise<unknown>;
  listEntityAttributes(entityLogicalName: string, attributeType?: string): Promise<{
    logicalName: string;
    displayName: string;
    attributeType: string;
    isRequired: boolean;
    isCustomAttribute: boolean;
    description?: string;
  }[]>;
  getAttributeMetadata(entityLogicalName: string, attributeLogicalName: string): Promise<unknown>;
  getOptionSetValues(entityLogicalName: string, attributeLogicalName: string): Promise<{ value: number; label: string }[]>;
  getGlobalOptionSet(optionSetName: string): Promise<{ value: number; label: string }[]>;

  // Relationships
  associateRecords(sourceEntityType: string, sourceId: string, targetEntityType: string, targetId: string, relationshipName: string): Promise<void>;
  disassociateRecords(sourceEntityType: string, sourceId: string, targetEntityType: string, targetId: string, relationshipName: string): Promise<void>;
  listRelatedRecords(entityType: string, id: string, navigationProperty: string, params?: { select?: string; limit?: number }): Promise<unknown[]>;
}

// =============================================================================
// Dynamics 365 Client Implementation
// =============================================================================

class DynamicsClientImpl implements CrmClient {
  private credentials: TenantCredentials;
  private baseUrl: string;
  private apiVersion = 'v9.2';

  // Token caching for client credentials flow
  private accessToken: string | null = null;
  private tokenExpiry = 0;

  constructor(credentials: TenantCredentials) {
    this.credentials = credentials;
    // Environment URL comes from X-Dynamics-Environment-URL header
    this.baseUrl = `${credentials.dynamicsEnvironmentUrl}/api/data/${this.apiVersion}`;
  }

  // ===========================================================================
  // Authentication
  // ===========================================================================

  private async getAccessToken(): Promise<string> {
    // If access token was passed directly via header, use it
    if (this.credentials.accessToken) {
      return this.credentials.accessToken;
    }

    // Check cached token
    if (this.accessToken && Date.now() < this.tokenExpiry) {
      return this.accessToken;
    }

    // Credentials come from X-Dynamics-Tenant-ID, X-CRM-Client-ID, X-CRM-Client-Secret headers
    if (
      !this.credentials.dynamicsTenantId ||
      !this.credentials.clientId ||
      !this.credentials.clientSecret
    ) {
      throw new AuthenticationError(
        'Missing credentials for OAuth client credentials flow. ' +
          'Provide X-Dynamics-Tenant-ID, X-CRM-Client-ID, and X-CRM-Client-Secret headers, ' +
          'or use X-CRM-Access-Token for pre-obtained tokens.'
      );
    }

    const tokenUrl = `https://login.microsoftonline.com/${this.credentials.dynamicsTenantId}/oauth2/v2.0/token`;

    const params = new URLSearchParams({
      grant_type: 'client_credentials',
      client_id: this.credentials.clientId,
      client_secret: this.credentials.clientSecret,
      scope: `${this.credentials.dynamicsEnvironmentUrl}/.default`,
    });

    const response = await fetch(tokenUrl, {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: params,
    });

    if (!response.ok) {
      const errorBody = await response.text();
      let errorMessage = 'Failed to acquire Dynamics 365 token';
      try {
        const errorJson = JSON.parse(errorBody);
        errorMessage = errorJson.error_description || errorJson.error || errorMessage;
      } catch {
        // Use default message
      }
      throw new AuthenticationError(errorMessage);
    }

    const data = (await response.json()) as {
      access_token: string;
      expires_in: number;
    };
    this.accessToken = data.access_token;
    // Expire 1 minute early to avoid edge cases
    this.tokenExpiry = Date.now() + data.expires_in * 1000 - 60000;
    return this.accessToken;
  }

  private async getAuthHeaders(): Promise<Record<string, string>> {
    const token = await this.getAccessToken();
    return {
      Authorization: `Bearer ${token}`,
      'Content-Type': 'application/json',
      'OData-MaxVersion': '4.0',
      'OData-Version': '4.0',
      Accept: 'application/json',
      Prefer: 'odata.include-annotations="*"',
    };
  }

  // ===========================================================================
  // HTTP Request Helper
  // ===========================================================================

  private async request<T>(
    endpoint: string,
    options: RequestInit = {}
  ): Promise<T> {
    const url = endpoint.startsWith('http') ? endpoint : `${this.baseUrl}${endpoint}`;

    const authHeaders = await this.getAuthHeaders();

    const response = await fetch(url, {
      ...options,
      headers: {
        ...authHeaders,
        ...(options.headers || {}),
      },
    });

    // Handle rate limiting
    if (response.status === 429) {
      const retryAfter = response.headers.get('Retry-After');
      throw new RateLimitError(
        'Dynamics 365 rate limit exceeded',
        retryAfter ? parseInt(retryAfter, 10) : 60
      );
    }

    // Handle authentication errors
    if (response.status === 401) {
      // Clear cached token and retry once
      this.accessToken = null;
      this.tokenExpiry = 0;
      throw new AuthenticationError('Authentication failed. Token may have expired.');
    }

    if (response.status === 403) {
      throw new AuthenticationError(
        'Access denied. Check that your app has the required Dynamics 365 permissions.'
      );
    }

    // Handle other errors
    if (!response.ok) {
      const errorBody = await response.text();
      let message = `Dynamics 365 API error: ${response.status}`;
      try {
        const errorJson = JSON.parse(errorBody);
        message = errorJson.error?.message || errorJson.message || message;
      } catch {
        // Use default message
      }
      throw new CrmApiError(message, response.status);
    }

    // Handle 204 No Content
    if (response.status === 204) {
      return undefined as T;
    }

    return response.json() as Promise<T>;
  }

  // ===========================================================================
  // Entity ID Extraction from OData-EntityId header
  // ===========================================================================

  private async createEntity<T>(
    endpoint: string,
    body: Record<string, unknown>
  ): Promise<{ id: string; response: Response }> {
    const url = `${this.baseUrl}${endpoint}`;
    const authHeaders = await this.getAuthHeaders();

    const response = await fetch(url, {
      method: 'POST',
      headers: {
        ...authHeaders,
        Prefer: 'return=representation',
      },
      body: JSON.stringify(body),
    });

    if (!response.ok) {
      const errorBody = await response.text();
      let message = `Dynamics 365 API error: ${response.status}`;
      try {
        const errorJson = JSON.parse(errorBody);
        message = errorJson.error?.message || errorJson.message || message;
      } catch {
        // Use default message
      }
      throw new CrmApiError(message, response.status);
    }

    // Extract entity ID from OData-EntityId header
    const entityIdHeader = response.headers.get('OData-EntityId');
    let id = '';
    if (entityIdHeader) {
      // Format: https://org.crm.dynamics.com/api/data/v9.2/contacts(00000000-0000-0000-0000-000000000001)
      const match = entityIdHeader.match(/\(([^)]+)\)$/);
      if (match) {
        id = match[1];
      }
    }

    return { id, response };
  }

  // ===========================================================================
  // Connection
  // ===========================================================================

  async testConnection(): Promise<{ connected: boolean; message: string }> {
    try {
      const whoAmI = await this.request<DynamicsWhoAmI>('/WhoAmI');
      return {
        connected: true,
        message: `Successfully connected to Dynamics 365. User ID: ${whoAmI.UserId}, Organization ID: ${whoAmI.OrganizationId}`,
      };
    } catch (error) {
      return {
        connected: false,
        message: error instanceof Error ? error.message : 'Connection failed',
      };
    }
  }

  // ===========================================================================
  // Contacts
  // ===========================================================================

  private mapContact(d: DynamicsContact): Contact {
    return {
      id: d.contactid,
      firstName: d.firstname,
      lastName: d.lastname,
      fullName: d.fullname,
      email: d.emailaddress1,
      phone: d.telephone1,
      mobilePhone: d.mobilephone,
      title: d.jobtitle,
      department: d.department,
      companyId: d._parentcustomerid_value,
      status: d.statecode === 0 ? 'active' : 'inactive',
      createdAt: d.createdon,
      updatedAt: d.modifiedon,
      ownerId: d._ownerid_value,
    };
  }

  async listContacts(params?: PaginationParams): Promise<PaginatedResponse<Contact>> {
    const limit = params?.limit || 20;
    const offset = params?.offset || 0;

    const queryParams = new URLSearchParams({
      $select:
        'contactid,firstname,lastname,fullname,emailaddress1,telephone1,mobilephone,jobtitle,department,_parentcustomerid_value,createdon,modifiedon,statecode,_ownerid_value',
      $top: String(limit),
      $skip: String(offset),
      $orderby: 'modifiedon desc',
    });

    // Use cursor (nextLink) if provided
    let url = `/contacts?${queryParams}`;
    if (params?.cursor) {
      url = params.cursor;
    }

    const response = await this.request<DynamicsODataResponse<DynamicsContact>>(url);

    return {
      items: response.value.map((c) => this.mapContact(c)),
      count: response.value.length,
      total: response['@odata.count'],
      hasMore: !!response['@odata.nextLink'],
      nextCursor: response['@odata.nextLink'],
    };
  }

  async getContact(id: string): Promise<Contact> {
    const queryParams = new URLSearchParams({
      $select:
        'contactid,firstname,lastname,fullname,emailaddress1,telephone1,mobilephone,jobtitle,department,_parentcustomerid_value,createdon,modifiedon,statecode,_ownerid_value',
    });

    const contact = await this.request<DynamicsContact>(`/contacts(${id})?${queryParams}`);
    return this.mapContact(contact);
  }

  async createContact(input: ContactCreateInput): Promise<Contact> {
    const body: Record<string, unknown> = {
      firstname: input.firstName,
      lastname: input.lastName,
      emailaddress1: input.email,
      telephone1: input.phone,
      jobtitle: input.title,
    };

    // Handle company association
    if (input.companyId) {
      body['parentcustomerid_account@odata.bind'] = `/accounts(${input.companyId})`;
    }

    // Add custom fields
    if (input.customFields) {
      Object.assign(body, input.customFields);
    }

    const { id } = await this.createEntity('/contacts', body);
    return this.getContact(id);
  }

  async updateContact(id: string, input: ContactUpdateInput): Promise<Contact> {
    const body: Record<string, unknown> = {};

    if (input.firstName !== undefined) body.firstname = input.firstName;
    if (input.lastName !== undefined) body.lastname = input.lastName;
    if (input.email !== undefined) body.emailaddress1 = input.email;
    if (input.phone !== undefined) body.telephone1 = input.phone;
    if (input.title !== undefined) body.jobtitle = input.title;

    // Handle company association
    if (input.companyId !== undefined) {
      if (input.companyId) {
        body['parentcustomerid_account@odata.bind'] = `/accounts(${input.companyId})`;
      } else {
        // To remove association, we need to use $ref endpoint
        body['parentcustomerid_account@odata.bind'] = null;
      }
    }

    // Add custom fields
    if (input.customFields) {
      Object.assign(body, input.customFields);
    }

    await this.request(`/contacts(${id})`, {
      method: 'PATCH',
      body: JSON.stringify(body),
    });

    return this.getContact(id);
  }

  async deleteContact(id: string): Promise<void> {
    await this.request(`/contacts(${id})`, {
      method: 'DELETE',
    });
  }

  async searchContacts(params: SearchParams): Promise<PaginatedResponse<Contact>> {
    const limit = params?.limit || 20;
    const offset = params?.offset || 0;
    const query = params.query || '';

    // Build filter using contains for fullname and email
    const filters: string[] = [];
    if (query) {
      filters.push(
        `(contains(fullname,'${query}') or contains(emailaddress1,'${query}'))`
      );
    }

    // Add additional filters
    if (params.filters) {
      for (const filter of params.filters) {
        const value =
          typeof filter.value === 'string' ? `'${filter.value}'` : filter.value;
        switch (filter.operator) {
          case 'eq':
            filters.push(`${filter.field} eq ${value}`);
            break;
          case 'contains':
            filters.push(`contains(${filter.field},${value})`);
            break;
          case 'starts_with':
            filters.push(`startswith(${filter.field},${value})`);
            break;
          default:
            // Other operators can be added as needed
            break;
        }
      }
    }

    const queryParams = new URLSearchParams({
      $select:
        'contactid,firstname,lastname,fullname,emailaddress1,telephone1,mobilephone,jobtitle,department,_parentcustomerid_value,createdon,modifiedon,statecode,_ownerid_value',
      $top: String(limit),
      $skip: String(offset),
    });

    if (filters.length > 0) {
      queryParams.set('$filter', filters.join(' and '));
    }

    if (params.sortBy) {
      queryParams.set('$orderby', `${params.sortBy} ${params.sortOrder || 'asc'}`);
    }

    const response = await this.request<DynamicsODataResponse<DynamicsContact>>(
      `/contacts?${queryParams}`
    );

    return {
      items: response.value.map((c) => this.mapContact(c)),
      count: response.value.length,
      total: response['@odata.count'],
      hasMore: !!response['@odata.nextLink'],
      nextCursor: response['@odata.nextLink'],
    };
  }

  // ===========================================================================
  // Companies (Accounts in Dynamics)
  // ===========================================================================

  private mapCompany(d: DynamicsAccount): Company {
    return {
      id: d.accountid,
      name: d.name,
      website: d.websiteurl,
      industry: d.industrycode?.toString(),
      description: d.description,
      numberOfEmployees: d.numberofemployees,
      annualRevenue: d.revenue,
      phone: d.telephone1,
      address: {
        street: d.address1_line1,
        city: d.address1_city,
        state: d.address1_stateorprovince,
        country: d.address1_country,
        postalCode: d.address1_postalcode,
      },
      createdAt: d.createdon,
      updatedAt: d.modifiedon,
      ownerId: d._ownerid_value,
    };
  }

  async listCompanies(params?: PaginationParams): Promise<PaginatedResponse<Company>> {
    const limit = params?.limit || 20;
    const offset = params?.offset || 0;

    const queryParams = new URLSearchParams({
      $select:
        'accountid,name,websiteurl,industrycode,description,numberofemployees,revenue,telephone1,address1_line1,address1_city,address1_stateorprovince,address1_country,address1_postalcode,createdon,modifiedon,statecode,_ownerid_value',
      $top: String(limit),
      $skip: String(offset),
      $orderby: 'modifiedon desc',
    });

    let url = `/accounts?${queryParams}`;
    if (params?.cursor) {
      url = params.cursor;
    }

    const response = await this.request<DynamicsODataResponse<DynamicsAccount>>(url);

    return {
      items: response.value.map((a) => this.mapCompany(a)),
      count: response.value.length,
      total: response['@odata.count'],
      hasMore: !!response['@odata.nextLink'],
      nextCursor: response['@odata.nextLink'],
    };
  }

  async getCompany(id: string): Promise<Company> {
    const queryParams = new URLSearchParams({
      $select:
        'accountid,name,websiteurl,industrycode,description,numberofemployees,revenue,telephone1,address1_line1,address1_city,address1_stateorprovince,address1_country,address1_postalcode,createdon,modifiedon,statecode,_ownerid_value',
    });

    const account = await this.request<DynamicsAccount>(`/accounts(${id})?${queryParams}`);
    return this.mapCompany(account);
  }

  async createCompany(input: CompanyCreateInput): Promise<Company> {
    const body: Record<string, unknown> = {
      name: input.name,
      websiteurl: input.domain,
      description: input.description,
      numberofemployees: input.numberOfEmployees,
      telephone1: input.phone,
    };

    // Industry is typically an OptionSet value in Dynamics
    if (input.industry) {
      // If it's a number, use directly; otherwise, you'd need a mapping
      const industryCode = parseInt(input.industry, 10);
      if (!isNaN(industryCode)) {
        body.industrycode = industryCode;
      }
    }

    // Address
    if (input.address) {
      if (input.address.street) body.address1_line1 = input.address.street;
      if (input.address.city) body.address1_city = input.address.city;
      if (input.address.state) body.address1_stateorprovince = input.address.state;
      if (input.address.country) body.address1_country = input.address.country;
      if (input.address.postalCode) body.address1_postalcode = input.address.postalCode;
    }

    // Custom fields
    if (input.customFields) {
      Object.assign(body, input.customFields);
    }

    const { id } = await this.createEntity('/accounts', body);
    return this.getCompany(id);
  }

  async updateCompany(id: string, input: CompanyUpdateInput): Promise<Company> {
    const body: Record<string, unknown> = {};

    if (input.name !== undefined) body.name = input.name;
    if (input.domain !== undefined) body.websiteurl = input.domain;
    if (input.description !== undefined) body.description = input.description;
    if (input.numberOfEmployees !== undefined)
      body.numberofemployees = input.numberOfEmployees;
    if (input.phone !== undefined) body.telephone1 = input.phone;

    if (input.industry !== undefined) {
      const industryCode = parseInt(input.industry, 10);
      if (!isNaN(industryCode)) {
        body.industrycode = industryCode;
      }
    }

    // Address
    if (input.address) {
      if (input.address.street !== undefined) body.address1_line1 = input.address.street;
      if (input.address.city !== undefined) body.address1_city = input.address.city;
      if (input.address.state !== undefined)
        body.address1_stateorprovince = input.address.state;
      if (input.address.country !== undefined) body.address1_country = input.address.country;
      if (input.address.postalCode !== undefined)
        body.address1_postalcode = input.address.postalCode;
    }

    // Custom fields
    if (input.customFields) {
      Object.assign(body, input.customFields);
    }

    await this.request(`/accounts(${id})`, {
      method: 'PATCH',
      body: JSON.stringify(body),
    });

    return this.getCompany(id);
  }

  // ===========================================================================
  // Deals (Opportunities in Dynamics)
  // ===========================================================================

  private mapDeal(d: DynamicsOpportunity): Deal {
    let status: 'open' | 'won' | 'lost' = 'open';
    if (d.statecode === 1) status = 'won';
    else if (d.statecode === 2) status = 'lost';

    return {
      id: d.opportunityid,
      name: d.name,
      amount: d.estimatedvalue,
      currency: d._transactioncurrencyid_value,
      stage: d.stepname,
      closeDate: d.estimatedclosedate,
      probability: d.closeprobability,
      companyId: d._parentaccountid_value,
      description: d.description,
      status,
      createdAt: d.createdon,
      updatedAt: d.modifiedon,
      ownerId: d._ownerid_value,
    };
  }

  async listDeals(params?: PaginationParams): Promise<PaginatedResponse<Deal>> {
    const limit = params?.limit || 20;
    const offset = params?.offset || 0;

    const queryParams = new URLSearchParams({
      $select:
        'opportunityid,name,estimatedvalue,_transactioncurrencyid_value,stepname,estimatedclosedate,closeprobability,_parentaccountid_value,description,statecode,statuscode,createdon,modifiedon,_ownerid_value',
      $top: String(limit),
      $skip: String(offset),
      $orderby: 'modifiedon desc',
    });

    let url = `/opportunities?${queryParams}`;
    if (params?.cursor) {
      url = params.cursor;
    }

    const response = await this.request<DynamicsODataResponse<DynamicsOpportunity>>(url);

    return {
      items: response.value.map((o) => this.mapDeal(o)),
      count: response.value.length,
      total: response['@odata.count'],
      hasMore: !!response['@odata.nextLink'],
      nextCursor: response['@odata.nextLink'],
    };
  }

  async getDeal(id: string): Promise<Deal> {
    const queryParams = new URLSearchParams({
      $select:
        'opportunityid,name,estimatedvalue,_transactioncurrencyid_value,stepname,estimatedclosedate,closeprobability,_parentaccountid_value,description,statecode,statuscode,createdon,modifiedon,_ownerid_value',
    });

    const opportunity = await this.request<DynamicsOpportunity>(
      `/opportunities(${id})?${queryParams}`
    );
    return this.mapDeal(opportunity);
  }

  async createDeal(input: DealCreateInput): Promise<Deal> {
    const body: Record<string, unknown> = {
      name: input.name,
      estimatedvalue: input.amount,
      estimatedclosedate: input.closeDate,
    };

    // Handle account association
    if (input.companyId) {
      body['parentaccountid@odata.bind'] = `/accounts(${input.companyId})`;
    }

    // Stage
    if (input.stageId) {
      body.stepname = input.stageId;
    }

    // Custom fields
    if (input.customFields) {
      Object.assign(body, input.customFields);
    }

    const { id } = await this.createEntity('/opportunities', body);
    return this.getDeal(id);
  }

  async updateDeal(id: string, input: DealUpdateInput): Promise<Deal> {
    const body: Record<string, unknown> = {};

    if (input.name !== undefined) body.name = input.name;
    if (input.amount !== undefined) body.estimatedvalue = input.amount;
    if (input.closeDate !== undefined) body.estimatedclosedate = input.closeDate;
    if (input.stageId !== undefined) body.stepname = input.stageId;

    // Status changes in Dynamics require specific state/status transitions
    // For simplicity, we handle basic field updates here
    if (input.status !== undefined) {
      // In Dynamics, changing status requires special handling
      // statecode: 0=Open, 1=Won, 2=Lost
      switch (input.status) {
        case 'won':
          body.statecode = 1;
          body.statuscode = 3; // Won
          break;
        case 'lost':
          body.statecode = 2;
          body.statuscode = 4; // Lost
          break;
        case 'open':
          body.statecode = 0;
          body.statuscode = 1; // In Progress
          break;
      }
    }

    // Custom fields
    if (input.customFields) {
      Object.assign(body, input.customFields);
    }

    await this.request(`/opportunities(${id})`, {
      method: 'PATCH',
      body: JSON.stringify(body),
    });

    return this.getDeal(id);
  }

  async moveDealStage(id: string, stageId: string): Promise<Deal> {
    // In Dynamics, stepname represents the current stage
    // For Business Process Flows, you'd need to update the active stage differently
    await this.request(`/opportunities(${id})`, {
      method: 'PATCH',
      body: JSON.stringify({ stepname: stageId }),
    });

    return this.getDeal(id);
  }

  async listPipelines(): Promise<Pipeline[]> {
    // In Dynamics 365, pipelines are Business Process Flows (workflows with category = 4)
    const workflowParams = new URLSearchParams({
      $filter: 'category eq 4 and statecode eq 1',
      $select: 'workflowid,name',
    });

    const workflows = await this.request<DynamicsODataResponse<DynamicsWorkflow>>(
      `/workflows?${workflowParams}`
    );

    const pipelines: Pipeline[] = [];

    for (const workflow of workflows.value) {
      // Get stages for this process
      const stagesParams = new URLSearchParams({
        $filter: `_processid_value eq ${workflow.workflowid}`,
        $select: 'processstageid,stagename,stagecategory',
        $orderby: 'stagecategory asc',
      });

      const stages = await this.request<DynamicsODataResponse<DynamicsProcessStage>>(
        `/processstages?${stagesParams}`
      );

      const pipelineStages: PipelineStage[] = stages.value.map((s, index) => ({
        id: s.processstageid,
        name: s.stagename,
        order: index,
        // stagecategory: 0=Qualify, 1=Develop, 2=Propose, 3=Close
        isClosed: s.stagecategory === 3,
        isWon: s.stagecategory === 3,
      }));

      pipelines.push({
        id: workflow.workflowid,
        name: workflow.name,
        stages: pipelineStages,
      });
    }

    return pipelines;
  }

  // ===========================================================================
  // Activities
  // ===========================================================================

  private mapActivity(d: DynamicsActivityPointer): Activity {
    let status: 'pending' | 'completed' | 'cancelled' = 'pending';
    if (d.statecode === 1) status = 'completed';
    else if (d.statecode === 2) status = 'cancelled';

    let type: 'call' | 'email' | 'meeting' | 'task' | 'note' | 'other' = 'other';
    switch (d.activitytypecode) {
      case 'phonecall':
        type = 'call';
        break;
      case 'email':
        type = 'email';
        break;
      case 'appointment':
        type = 'meeting';
        break;
      case 'task':
        type = 'task';
        break;
      case 'annotation':
        type = 'note';
        break;
    }

    return {
      id: d.activityid,
      type,
      subject: d.subject || '',
      body: d.description,
      status,
      dueDate: d.scheduledend,
      completedDate: d.actualend,
      durationMinutes: d.actualdurationminutes,
      activityDate: d.scheduledstart,
      createdAt: d.createdon,
      updatedAt: d.modifiedon,
      ownerId: d._ownerid_value,
    };
  }

  async listActivities(
    params?: PaginationParams & { recordId?: string }
  ): Promise<PaginatedResponse<Activity>> {
    const limit = params?.limit || 20;
    const offset = params?.offset || 0;

    const queryParams = new URLSearchParams({
      $select:
        'activityid,subject,description,activitytypecode,scheduledstart,scheduledend,actualend,actualdurationminutes,statecode,statuscode,createdon,modifiedon,_regardingobjectid_value,_ownerid_value',
      $top: String(limit),
      $skip: String(offset),
      $orderby: 'modifiedon desc',
    });

    // Filter by regarding object if provided
    if (params?.recordId) {
      queryParams.set('$filter', `_regardingobjectid_value eq ${params.recordId}`);
    }

    let url = `/activitypointers?${queryParams}`;
    if (params?.cursor) {
      url = params.cursor;
    }

    const response = await this.request<DynamicsODataResponse<DynamicsActivityPointer>>(url);

    return {
      items: response.value.map((a) => this.mapActivity(a)),
      count: response.value.length,
      total: response['@odata.count'],
      hasMore: !!response['@odata.nextLink'],
      nextCursor: response['@odata.nextLink'],
    };
  }

  async createActivity(input: ActivityCreateInput): Promise<Activity> {
    // Map activity type to Dynamics entity
    let endpoint = '/tasks';
    const body: Record<string, unknown> = {
      subject: input.subject,
      description: input.body,
      scheduledend: input.dueDate,
    };

    // Handle contact association
    if (input.contactIds && input.contactIds.length > 0) {
      body['regardingobjectid_contact@odata.bind'] = `/contacts(${input.contactIds[0]})`;
    } else if (input.companyId) {
      body['regardingobjectid_account@odata.bind'] = `/accounts(${input.companyId})`;
    } else if (input.dealId) {
      body['regardingobjectid_opportunity@odata.bind'] = `/opportunities(${input.dealId})`;
    }

    // Custom fields
    if (input.customFields) {
      Object.assign(body, input.customFields);
    }

    const { id } = await this.createEntity(endpoint, body);

    // Fetch the created activity
    const activity = await this.request<DynamicsActivityPointer>(
      `/tasks(${id})?$select=activityid,subject,description,activitytypecode,scheduledstart,scheduledend,actualend,actualdurationminutes,statecode,statuscode,createdon,modifiedon,_regardingobjectid_value,_ownerid_value`
    );

    return this.mapActivity(activity);
  }

  async logCall(
    contactId: string,
    subject: string,
    notes?: string,
    durationMinutes?: number
  ): Promise<Activity> {
    const body: Record<string, unknown> = {
      subject,
      description: notes,
      actualdurationminutes: durationMinutes || 0,
      phonecall_activity_parties: [
        {
          participationtypemask: 2, // To
          'partyid_contact@odata.bind': `/contacts(${contactId})`,
        },
      ],
      statecode: 1, // Completed
      statuscode: 2, // Made
    };

    const { id } = await this.createEntity('/phonecalls', body);

    const activity = await this.request<DynamicsActivityPointer>(
      `/phonecalls(${id})?$select=activityid,subject,description,activitytypecode,scheduledstart,scheduledend,actualend,actualdurationminutes,statecode,statuscode,createdon,modifiedon,_regardingobjectid_value,_ownerid_value`
    );

    return this.mapActivity(activity);
  }

  async logEmail(
    contactId: string,
    subject: string,
    body: string,
    direction: 'sent' | 'received'
  ): Promise<Activity> {
    const emailBody: Record<string, unknown> = {
      subject,
      description: body,
      directioncode: direction === 'sent', // true = outgoing, false = incoming
      email_activity_parties: [
        {
          participationtypemask: direction === 'sent' ? 2 : 1, // 1=From, 2=To
          'partyid_contact@odata.bind': `/contacts(${contactId})`,
        },
      ],
      statecode: 1, // Completed
      statuscode: direction === 'sent' ? 2 : 3, // 2=Sent, 3=Received
    };

    const { id } = await this.createEntity('/emails', emailBody);

    const activity = await this.request<DynamicsActivityPointer>(
      `/emails(${id})?$select=activityid,subject,description,activitytypecode,scheduledstart,scheduledend,actualend,actualdurationminutes,statecode,statuscode,createdon,modifiedon,_regardingobjectid_value,_ownerid_value`
    );

    return this.mapActivity(activity);
  }

  async getActivity(id: string, activityType?: string): Promise<Activity> {
    const entitySet = activityType ? `${activityType}s` : 'activitypointers';
    const activity = await this.request<DynamicsActivityPointer>(
      `/${entitySet}(${id})?$select=activityid,subject,description,activitytypecode,scheduledstart,scheduledend,actualend,actualdurationminutes,statecode,statuscode,createdon,modifiedon,_regardingobjectid_value,_ownerid_value`
    );
    return this.mapActivity(activity);
  }

  async updateActivity(id: string, activityType: string, input: { subject?: string; description?: string; scheduledStart?: string; scheduledEnd?: string; priorityCode?: number }): Promise<Activity> {
    const body: Record<string, unknown> = {};
    if (input.subject !== undefined) body.subject = input.subject;
    if (input.description !== undefined) body.description = input.description;
    if (input.scheduledStart !== undefined) body.scheduledstart = input.scheduledStart;
    if (input.scheduledEnd !== undefined) body.scheduledend = input.scheduledEnd;
    if (input.priorityCode !== undefined) body.prioritycode = input.priorityCode;
    await this.request(`/${activityType}s(${id})`, { method: 'PATCH', body: JSON.stringify(body) });
    return this.getActivity(id, activityType);
  }

  async deleteActivity(id: string, activityType: string): Promise<void> {
    await this.request(`/${activityType}s(${id})`, { method: 'DELETE' });
  }

  async completeActivity(id: string, activityType: string): Promise<void> {
    await this.request(`/${activityType}s(${id})`, {
      method: 'PATCH',
      body: JSON.stringify({ statecode: 1, statuscode: 2 }), // Completed
    });
  }

  async cancelActivity(id: string, activityType: string): Promise<void> {
    await this.request(`/${activityType}s(${id})`, {
      method: 'PATCH',
      body: JSON.stringify({ statecode: 2, statuscode: 3 }), // Canceled
    });
  }

  async createAppointment(input: {
    subject: string;
    description?: string;
    location?: string;
    scheduledStart: string;
    scheduledEnd: string;
    regardingId?: string;
    regardingType?: string;
    requiredAttendees?: string[];
    optionalAttendees?: string[];
    isAllDayEvent?: boolean;
  }): Promise<Activity> {
    const body: Record<string, unknown> = {
      subject: input.subject,
      description: input.description,
      location: input.location,
      scheduledstart: input.scheduledStart,
      scheduledend: input.scheduledEnd,
      isalldayevent: input.isAllDayEvent || false,
    };
    if (input.regardingId && input.regardingType) {
      body[`regardingobjectid_${input.regardingType}@odata.bind`] = `/${input.regardingType}s(${input.regardingId})`;
    }
    // Add required attendees as activity parties
    if (input.requiredAttendees?.length) {
      body.appointment_activity_parties = input.requiredAttendees.map(id => ({
        participationtypemask: 5, // Required Attendee
        'partyid_systemuser@odata.bind': `/systemusers(${id})`,
      }));
    }
    const { id } = await this.createEntity('/appointments', body);
    return this.getActivity(id, 'appointment');
  }

  async createLetter(input: {
    subject: string;
    description?: string;
    regardingId?: string;
    regardingType?: string;
    address?: string;
  }): Promise<Activity> {
    const body: Record<string, unknown> = {
      subject: input.subject,
      description: input.description,
      address: input.address,
    };
    if (input.regardingId && input.regardingType) {
      body[`regardingobjectid_${input.regardingType}@odata.bind`] = `/${input.regardingType}s(${input.regardingId})`;
    }
    const { id } = await this.createEntity('/letters', body);
    return this.getActivity(id, 'letter');
  }

  async createFax(input: {
    subject: string;
    description?: string;
    regardingId?: string;
    regardingType?: string;
    faxNumber?: string;
  }): Promise<Activity> {
    const body: Record<string, unknown> = {
      subject: input.subject,
      description: input.description,
      faxnumber: input.faxNumber,
    };
    if (input.regardingId && input.regardingType) {
      body[`regardingobjectid_${input.regardingType}@odata.bind`] = `/${input.regardingType}s(${input.regardingId})`;
    }
    const { id } = await this.createEntity('/faxes', body);
    return this.getActivity(id, 'fax');
  }

  async listActivitiesByType(activityType: string, params?: PaginationParams & { regardingId?: string }): Promise<PaginatedResponse<Activity>> {
    const limit = params?.limit || 20;
    const offset = params?.offset || 0;
    const entitySet = `${activityType}s`;
    const queryParams = new URLSearchParams({
      $select: 'activityid,subject,description,activitytypecode,scheduledstart,scheduledend,actualend,actualdurationminutes,statecode,statuscode,createdon,modifiedon,_regardingobjectid_value,_ownerid_value',
      $top: String(limit),
      $skip: String(offset),
      $orderby: 'createdon desc',
    });
    if (params?.regardingId) {
      queryParams.set('$filter', `_regardingobjectid_value eq ${params.regardingId}`);
    }
    let url = `/${entitySet}?${queryParams}`;
    if (params?.cursor) url = params.cursor;
    const response = await this.request<DynamicsODataResponse<DynamicsActivityPointer>>(url);
    return {
      items: response.value.map((a) => this.mapActivity(a)),
      count: response.value.length,
      total: response['@odata.count'],
      hasMore: !!response['@odata.nextLink'],
      nextCursor: response['@odata.nextLink'],
    };
  }

  async getActivityParties(activityId: string): Promise<unknown[]> {
    const response = await this.request<DynamicsODataResponse<unknown>>(
      `/activityparties?$filter=_activityid_value eq ${activityId}&$select=activitypartyid,participationtypemask,_partyid_value,addressused`
    );
    return response.value;
  }

  // ===========================================================================
  // Leads
  // ===========================================================================

  private mapLead(d: DynamicsLead): Lead {
    let status: 'open' | 'qualified' | 'disqualified' = 'open';
    if (d.statecode === 1) status = 'qualified';
    else if (d.statecode === 2) status = 'disqualified';

    return {
      id: d.leadid,
      subject: d.subject,
      firstName: d.firstname,
      lastName: d.lastname,
      fullName: d.fullname,
      email: d.emailaddress1,
      phone: d.telephone1,
      mobilePhone: d.mobilephone,
      companyName: d.companyname,
      jobTitle: d.jobtitle,
      website: d.websiteurl,
      description: d.description,
      address: {
        street: d.address1_line1,
        city: d.address1_city,
        state: d.address1_stateorprovince,
        postalCode: d.address1_postalcode,
        country: d.address1_country,
      },
      leadSource: d.leadsourcecode,
      leadQuality: d.leadqualitycode,
      industryCode: d.industrycode,
      revenue: d.revenue,
      numberOfEmployees: d.numberofemployees,
      status,
      stateCode: d.statecode,
      ownerId: d._ownerid_value,
      parentAccountId: d._parentaccountid_value,
      parentContactId: d._parentcontactid_value,
      createdAt: d.createdon,
      updatedAt: d.modifiedon,
    };
  }

  private readonly leadSelectFields =
    'leadid,subject,firstname,lastname,fullname,emailaddress1,telephone1,mobilephone,companyname,jobtitle,websiteurl,address1_line1,address1_city,address1_stateorprovince,address1_postalcode,address1_country,description,leadsourcecode,leadqualitycode,industrycode,revenue,numberofemployees,statuscode,statecode,_ownerid_value,_parentaccountid_value,_parentcontactid_value,createdon,modifiedon';

  async listLeads(params?: PaginationParams): Promise<PaginatedResponse<Lead>> {
    const limit = params?.limit || 20;
    const offset = params?.offset || 0;

    const queryParams = new URLSearchParams({
      $select: this.leadSelectFields,
      $top: String(limit),
      $skip: String(offset),
      $orderby: 'createdon desc',
    });

    let url = `/leads?${queryParams}`;
    if (params?.cursor) {
      url = params.cursor;
    }

    const response = await this.request<DynamicsODataResponse<DynamicsLead>>(url);

    return {
      items: response.value.map((l) => this.mapLead(l)),
      count: response.value.length,
      total: response['@odata.count'],
      hasMore: !!response['@odata.nextLink'],
      nextCursor: response['@odata.nextLink'],
    };
  }

  async getLead(id: string): Promise<Lead> {
    const queryParams = new URLSearchParams({
      $select: this.leadSelectFields,
    });

    const lead = await this.request<DynamicsLead>(`/leads(${id})?${queryParams}`);
    return this.mapLead(lead);
  }

  async createLead(input: LeadCreateInput): Promise<Lead> {
    const body: Record<string, unknown> = {
      subject: input.subject,
      firstname: input.firstName,
      lastname: input.lastName,
      emailaddress1: input.email,
      telephone1: input.phone,
      companyname: input.companyName,
      jobtitle: input.jobTitle,
      websiteurl: input.website,
      description: input.description,
      leadsourcecode: input.leadSource,
      leadqualitycode: input.leadQuality,
    };

    // Handle owner association
    if (input.ownerId) {
      body['ownerid@odata.bind'] = `/systemusers(${input.ownerId})`;
    }

    // Add custom fields
    if (input.customFields) {
      Object.assign(body, input.customFields);
    }

    const { id } = await this.createEntity('/leads', body);
    return this.getLead(id);
  }

  async updateLead(id: string, input: LeadUpdateInput): Promise<Lead> {
    const body: Record<string, unknown> = {};

    if (input.subject !== undefined) body.subject = input.subject;
    if (input.firstName !== undefined) body.firstname = input.firstName;
    if (input.lastName !== undefined) body.lastname = input.lastName;
    if (input.email !== undefined) body.emailaddress1 = input.email;
    if (input.phone !== undefined) body.telephone1 = input.phone;
    if (input.companyName !== undefined) body.companyname = input.companyName;
    if (input.jobTitle !== undefined) body.jobtitle = input.jobTitle;
    if (input.website !== undefined) body.websiteurl = input.website;
    if (input.description !== undefined) body.description = input.description;
    if (input.leadSource !== undefined) body.leadsourcecode = input.leadSource;
    if (input.leadQuality !== undefined) body.leadqualitycode = input.leadQuality;

    // Add custom fields
    if (input.customFields) {
      Object.assign(body, input.customFields);
    }

    await this.request(`/leads(${id})`, {
      method: 'PATCH',
      body: JSON.stringify(body),
    });

    return this.getLead(id);
  }

  async deleteLead(id: string): Promise<void> {
    await this.request(`/leads(${id})`, {
      method: 'DELETE',
    });
  }

  async searchLeads(params: SearchParams): Promise<PaginatedResponse<Lead>> {
    const limit = params?.limit || 20;
    const offset = params?.offset || 0;
    const query = params.query || '';

    // Build filter using contains for fullname, email, and company name
    const filters: string[] = [];
    if (query) {
      filters.push(
        `(contains(fullname,'${query}') or contains(emailaddress1,'${query}') or contains(companyname,'${query}'))`
      );
    }

    // Add additional filters
    if (params.filters) {
      for (const filter of params.filters) {
        const value =
          typeof filter.value === 'string' ? `'${filter.value}'` : filter.value;
        switch (filter.operator) {
          case 'eq':
            filters.push(`${filter.field} eq ${value}`);
            break;
          case 'contains':
            filters.push(`contains(${filter.field},${value})`);
            break;
          case 'starts_with':
            filters.push(`startswith(${filter.field},${value})`);
            break;
          default:
            break;
        }
      }
    }

    const queryParams = new URLSearchParams({
      $select: this.leadSelectFields,
      $top: String(limit),
      $skip: String(offset),
    });

    if (filters.length > 0) {
      queryParams.set('$filter', filters.join(' and '));
    }

    if (params.sortBy) {
      queryParams.set('$orderby', `${params.sortBy} ${params.sortOrder || 'asc'}`);
    }

    const response = await this.request<DynamicsODataResponse<DynamicsLead>>(
      `/leads?${queryParams}`
    );

    return {
      items: response.value.map((l) => this.mapLead(l)),
      count: response.value.length,
      total: response['@odata.count'],
      hasMore: !!response['@odata.nextLink'],
      nextCursor: response['@odata.nextLink'],
    };
  }

  async qualifyLead(id: string, input: LeadQualifyInput): Promise<LeadQualifyResult> {
    // Dynamics 365 QualifyLead action
    const body: Record<string, unknown> = {
      CreateAccount: input.createAccount,
      CreateContact: input.createContact,
      CreateOpportunity: input.createOpportunity,
      Status: input.status || 3, // 3 = Qualified
    };

    if (input.opportunityCurrencyId) {
      body.OpportunityCurrencyId = { '@odata.type': 'Microsoft.Dynamics.CRM.transactioncurrency', transactioncurrencyid: input.opportunityCurrencyId };
    }

    if (input.opportunityCustomerId) {
      body.OpportunityCustomerId = { '@odata.type': 'Microsoft.Dynamics.CRM.account', accountid: input.opportunityCustomerId };
    }

    if (input.sourceCampaignId) {
      body.SourceCampaignId = { '@odata.type': 'Microsoft.Dynamics.CRM.campaign', campaignid: input.sourceCampaignId };
    }

    const response = await this.request<{ CreatedEntities: Array<{ '@odata.type': string; [key: string]: unknown }> }>(
      `/leads(${id})/Microsoft.Dynamics.CRM.QualifyLead`,
      {
        method: 'POST',
        body: JSON.stringify(body),
      }
    );

    const result: LeadQualifyResult = { success: true };

    // Parse created entities from response
    if (response.CreatedEntities) {
      for (const entity of response.CreatedEntities) {
        const type = entity['@odata.type'];
        if (type?.includes('account')) {
          result.accountId = entity.accountid as string;
        } else if (type?.includes('contact')) {
          result.contactId = entity.contactid as string;
        } else if (type?.includes('opportunity')) {
          result.opportunityId = entity.opportunityid as string;
        }
      }
    }

    return result;
  }

  // ===========================================================================
  // Quotes
  // ===========================================================================

  private mapQuote(d: DynamicsQuote): Quote {
    let status: QuoteStatus = 'draft';
    if (d.statecode === 1) status = 'active';
    else if (d.statecode === 2) status = 'won';
    else if (d.statecode === 3) status = 'closed';

    return {
      id: d.quoteid,
      name: d.name,
      quoteNumber: d.quotenumber,
      description: d.description,
      effectiveFrom: d.effectivefrom,
      effectiveTo: d.effectiveto,
      totalAmount: d.totalamount,
      totalLineItemAmount: d.totallineitemamount,
      totalDiscountAmount: d.totaldiscountamount,
      totalTax: d.totaltax,
      freightAmount: d.freightamount,
      discountPercentage: d.discountpercentage,
      discountAmount: d.discountamount,
      status,
      stateCode: d.statecode,
      opportunityId: d._opportunityid_value,
      customerId: d._customerid_value,
      priceLevelId: d._pricelevelid_value,
      ownerId: d._ownerid_value,
      currencyId: d._transactioncurrencyid_value,
      createdAt: d.createdon,
      updatedAt: d.modifiedon,
    };
  }

  private readonly quoteSelectFields =
    'quoteid,name,quotenumber,description,effectivefrom,effectiveto,totalamount,totallineitemamount,totaldiscountamount,totaltax,freightamount,discountpercentage,discountamount,statuscode,statecode,_opportunityid_value,_customerid_value,_pricelevelid_value,_ownerid_value,_transactioncurrencyid_value,createdon,modifiedon';

  async listQuotes(params?: PaginationParams): Promise<PaginatedResponse<Quote>> {
    const limit = params?.limit || 20;
    const offset = params?.offset || 0;
    const queryParams = new URLSearchParams({
      $select: this.quoteSelectFields,
      $top: String(limit),
      $skip: String(offset),
      $orderby: 'createdon desc',
    });
    let url = `/quotes?${queryParams}`;
    if (params?.cursor) url = params.cursor;
    const response = await this.request<DynamicsODataResponse<DynamicsQuote>>(url);
    return {
      items: response.value.map((q) => this.mapQuote(q)),
      count: response.value.length,
      total: response['@odata.count'],
      hasMore: !!response['@odata.nextLink'],
      nextCursor: response['@odata.nextLink'],
    };
  }

  async getQuote(id: string): Promise<Quote> {
    const quote = await this.request<DynamicsQuote>(`/quotes(${id})?$select=${this.quoteSelectFields}`);
    return this.mapQuote(quote);
  }

  async createQuote(input: QuoteCreateInput): Promise<Quote> {
    const body: Record<string, unknown> = { name: input.name, description: input.description };
    if (input.opportunityId) body['opportunityid@odata.bind'] = `/opportunities(${input.opportunityId})`;
    if (input.customerAccountId) body['customerid_account@odata.bind'] = `/accounts(${input.customerAccountId})`;
    if (input.priceLevelId) body['pricelevelid@odata.bind'] = `/pricelevels(${input.priceLevelId})`;
    if (input.currencyId) body['transactioncurrencyid@odata.bind'] = `/transactioncurrencies(${input.currencyId})`;
    if (input.effectiveFrom) body.effectivefrom = input.effectiveFrom;
    if (input.effectiveTo) body.effectiveto = input.effectiveTo;
    if (input.discountPercentage) body.discountpercentage = input.discountPercentage;
    if (input.freightAmount) body.freightamount = input.freightAmount;
    if (input.customFields) Object.assign(body, input.customFields);
    const { id } = await this.createEntity('/quotes', body);
    return this.getQuote(id);
  }

  async updateQuote(id: string, input: Partial<QuoteCreateInput>): Promise<Quote> {
    const body: Record<string, unknown> = {};
    if (input.name !== undefined) body.name = input.name;
    if (input.description !== undefined) body.description = input.description;
    if (input.effectiveFrom !== undefined) body.effectivefrom = input.effectiveFrom;
    if (input.effectiveTo !== undefined) body.effectiveto = input.effectiveTo;
    if (input.discountPercentage !== undefined) body.discountpercentage = input.discountPercentage;
    if (input.freightAmount !== undefined) body.freightamount = input.freightAmount;
    if (input.customFields) Object.assign(body, input.customFields);
    await this.request(`/quotes(${id})`, { method: 'PATCH', body: JSON.stringify(body) });
    return this.getQuote(id);
  }

  async deleteQuote(id: string): Promise<void> {
    await this.request(`/quotes(${id})`, { method: 'DELETE' });
  }

  async listQuoteDetails(quoteId: string): Promise<QuoteDetail[]> {
    const response = await this.request<DynamicsODataResponse<DynamicsQuoteDetail>>(
      `/quotedetails?$filter=_quoteid_value eq ${quoteId}`
    );
    return response.value.map((d) => ({
      id: d.quotedetailid,
      quoteId: d._quoteid_value,
      productId: d._productid_value,
      productDescription: d.productdescription,
      quantity: d.quantity,
      pricePerUnit: d.priceperunit,
      baseAmount: d.baseamount,
      extendedAmount: d.extendedamount,
      manualDiscountAmount: d.manualdiscountamount,
      tax: d.tax,
      uomId: d._uomid_value,
      isPriceOverridden: d.ispriceoverridden,
      isProductOverridden: d.isproductoverridden,
    }));
  }

  async addQuoteDetail(input: QuoteDetailCreateInput): Promise<QuoteDetail> {
    const body: Record<string, unknown> = {
      'quoteid@odata.bind': `/quotes(${input.quoteId})`,
      quantity: input.quantity,
    };
    if (input.productId) body['productid@odata.bind'] = `/products(${input.productId})`;
    if (input.productDescription) body.productdescription = input.productDescription;
    if (input.pricePerUnit) body.priceperunit = input.pricePerUnit;
    if (input.manualDiscountAmount) body.manualdiscountamount = input.manualDiscountAmount;
    if (input.tax) body.tax = input.tax;
    if (input.uomId) body['uomid@odata.bind'] = `/uoms(${input.uomId})`;
    if (input.isPriceOverridden !== undefined) body.ispriceoverridden = input.isPriceOverridden;
    const { id } = await this.createEntity('/quotedetails', body);
    const detail = await this.request<DynamicsQuoteDetail>(`/quotedetails(${id})`);
    return {
      id: detail.quotedetailid,
      quoteId: detail._quoteid_value,
      productId: detail._productid_value,
      productDescription: detail.productdescription,
      quantity: detail.quantity,
      pricePerUnit: detail.priceperunit,
      baseAmount: detail.baseamount,
      extendedAmount: detail.extendedamount,
      manualDiscountAmount: detail.manualdiscountamount,
      tax: detail.tax,
      uomId: detail._uomid_value,
      isPriceOverridden: detail.ispriceoverridden,
      isProductOverridden: detail.isproductoverridden,
    };
  }

  async activateQuote(id: string): Promise<void> {
    await this.request(`/quotes(${id})/Microsoft.Dynamics.CRM.ActivateQuote`, { method: 'POST', body: '{}' });
  }

  async closeQuote(id: string, status: 'won' | 'lost' | 'cancelled'): Promise<void> {
    const statusMap = { won: 4, lost: 5, cancelled: 6 };
    await this.request(`/quotes(${id})/Microsoft.Dynamics.CRM.CloseQuote`, {
      method: 'POST',
      body: JSON.stringify({ Status: statusMap[status] }),
    });
  }

  async convertQuoteToOrder(quoteId: string): Promise<{ salesOrderId: string }> {
    const response = await this.request<{ salesorderid: string }>(
      `/quotes(${quoteId})/Microsoft.Dynamics.CRM.ConvertQuoteToSalesOrder`,
      { method: 'POST', body: JSON.stringify({ ColumnSet: { AllColumns: true } }) }
    );
    return { salesOrderId: response.salesorderid };
  }

  // ===========================================================================
  // Sales Orders
  // ===========================================================================

  private mapOrder(d: DynamicsSalesOrder): SalesOrder {
    let status: OrderStatus = 'active';
    if (d.statecode === 1) status = 'submitted';
    else if (d.statecode === 2) status = 'cancelled';
    else if (d.statecode === 3) status = 'fulfilled';
    else if (d.statecode === 4) status = 'invoiced';

    return {
      id: d.salesorderid,
      name: d.name,
      orderNumber: d.ordernumber,
      description: d.description,
      totalAmount: d.totalamount,
      totalLineItemAmount: d.totallineitemamount,
      totalDiscountAmount: d.totaldiscountamount,
      totalTax: d.totaltax,
      freightAmount: d.freightamount,
      discountPercentage: d.discountpercentage,
      discountAmount: d.discountamount,
      dateDelivered: d.datedelivered,
      requestDeliveryBy: d.requestdeliveryby,
      status,
      stateCode: d.statecode,
      quoteId: d._quoteid_value,
      opportunityId: d._opportunityid_value,
      customerId: d._customerid_value,
      priceLevelId: d._pricelevelid_value,
      ownerId: d._ownerid_value,
      createdAt: d.createdon,
      updatedAt: d.modifiedon,
    };
  }

  private readonly orderSelectFields =
    'salesorderid,name,ordernumber,description,totalamount,totallineitemamount,totaldiscountamount,totaltax,freightamount,discountpercentage,discountamount,datedelivered,requestdeliveryby,statuscode,statecode,_quoteid_value,_opportunityid_value,_customerid_value,_pricelevelid_value,_ownerid_value,createdon,modifiedon';

  async listOrders(params?: PaginationParams): Promise<PaginatedResponse<SalesOrder>> {
    const limit = params?.limit || 20;
    const offset = params?.offset || 0;
    const queryParams = new URLSearchParams({
      $select: this.orderSelectFields,
      $top: String(limit),
      $skip: String(offset),
      $orderby: 'createdon desc',
    });
    let url = `/salesorders?${queryParams}`;
    if (params?.cursor) url = params.cursor;
    const response = await this.request<DynamicsODataResponse<DynamicsSalesOrder>>(url);
    return {
      items: response.value.map((o) => this.mapOrder(o)),
      count: response.value.length,
      total: response['@odata.count'],
      hasMore: !!response['@odata.nextLink'],
      nextCursor: response['@odata.nextLink'],
    };
  }

  async getOrder(id: string): Promise<SalesOrder> {
    const order = await this.request<DynamicsSalesOrder>(`/salesorders(${id})?$select=${this.orderSelectFields}`);
    return this.mapOrder(order);
  }

  async createOrder(input: SalesOrderCreateInput): Promise<SalesOrder> {
    const body: Record<string, unknown> = { name: input.name, description: input.description };
    if (input.quoteId) body['quoteid@odata.bind'] = `/quotes(${input.quoteId})`;
    if (input.opportunityId) body['opportunityid@odata.bind'] = `/opportunities(${input.opportunityId})`;
    if (input.customerAccountId) body['customerid_account@odata.bind'] = `/accounts(${input.customerAccountId})`;
    if (input.priceLevelId) body['pricelevelid@odata.bind'] = `/pricelevels(${input.priceLevelId})`;
    if (input.requestDeliveryBy) body.requestdeliveryby = input.requestDeliveryBy;
    if (input.customFields) Object.assign(body, input.customFields);
    const { id } = await this.createEntity('/salesorders', body);
    return this.getOrder(id);
  }

  async updateOrder(id: string, input: Partial<SalesOrderCreateInput>): Promise<SalesOrder> {
    const body: Record<string, unknown> = {};
    if (input.name !== undefined) body.name = input.name;
    if (input.description !== undefined) body.description = input.description;
    if (input.requestDeliveryBy !== undefined) body.requestdeliveryby = input.requestDeliveryBy;
    if (input.customFields) Object.assign(body, input.customFields);
    await this.request(`/salesorders(${id})`, { method: 'PATCH', body: JSON.stringify(body) });
    return this.getOrder(id);
  }

  async deleteOrder(id: string): Promise<void> {
    await this.request(`/salesorders(${id})`, { method: 'DELETE' });
  }

  async listOrderDetails(orderId: string): Promise<SalesOrderDetail[]> {
    const response = await this.request<DynamicsODataResponse<{ salesorderdetailid: string; _salesorderid_value: string; _productid_value?: string; productdescription?: string; quantity?: number; priceperunit?: number; baseamount?: number; extendedamount?: number; manualdiscountamount?: number; tax?: number }>>(
      `/salesorderdetails?$filter=_salesorderid_value eq ${orderId}`
    );
    return response.value.map((d) => ({
      id: d.salesorderdetailid,
      salesOrderId: d._salesorderid_value,
      productId: d._productid_value,
      productDescription: d.productdescription,
      quantity: d.quantity,
      pricePerUnit: d.priceperunit,
      baseAmount: d.baseamount,
      extendedAmount: d.extendedamount,
      manualDiscountAmount: d.manualdiscountamount,
      tax: d.tax,
    }));
  }

  async addOrderDetail(input: SalesOrderDetailCreateInput): Promise<SalesOrderDetail> {
    const body: Record<string, unknown> = {
      'salesorderid@odata.bind': `/salesorders(${input.salesOrderId})`,
      quantity: input.quantity,
    };
    if (input.productId) body['productid@odata.bind'] = `/products(${input.productId})`;
    if (input.productDescription) body.productdescription = input.productDescription;
    if (input.pricePerUnit) body.priceperunit = input.pricePerUnit;
    if (input.manualDiscountAmount) body.manualdiscountamount = input.manualDiscountAmount;
    if (input.tax) body.tax = input.tax;
    const { id } = await this.createEntity('/salesorderdetails', body);
    const detail = await this.request<{ salesorderdetailid: string; _salesorderid_value: string; _productid_value?: string; productdescription?: string; quantity?: number; priceperunit?: number; baseamount?: number; extendedamount?: number; manualdiscountamount?: number; tax?: number }>(
      `/salesorderdetails(${id})`
    );
    return {
      id: detail.salesorderdetailid,
      salesOrderId: detail._salesorderid_value,
      productId: detail._productid_value,
      productDescription: detail.productdescription,
      quantity: detail.quantity,
      pricePerUnit: detail.priceperunit,
      baseAmount: detail.baseamount,
      extendedAmount: detail.extendedamount,
      manualDiscountAmount: detail.manualdiscountamount,
      tax: detail.tax,
    };
  }

  async fulfillOrder(id: string): Promise<void> {
    await this.request(`/salesorders(${id})/Microsoft.Dynamics.CRM.FulfillSalesOrder`, {
      method: 'POST',
      body: JSON.stringify({ OrderClose: { subject: 'Order fulfilled' }, Status: -1 }),
    });
  }

  async cancelOrder(id: string): Promise<void> {
    await this.request(`/salesorders(${id})/Microsoft.Dynamics.CRM.CancelSalesOrder`, {
      method: 'POST',
      body: JSON.stringify({ OrderClose: { subject: 'Order cancelled' }, Status: -1 }),
    });
  }

  async convertOrderToInvoice(orderId: string): Promise<{ invoiceId: string }> {
    const response = await this.request<{ invoiceid: string }>(
      `/salesorders(${orderId})/Microsoft.Dynamics.CRM.ConvertSalesOrderToInvoice`,
      { method: 'POST', body: JSON.stringify({ ColumnSet: { AllColumns: true } }) }
    );
    return { invoiceId: response.invoiceid };
  }

  // ===========================================================================
  // Invoices
  // ===========================================================================

  private mapInvoice(d: DynamicsInvoice): Invoice {
    let status: InvoiceStatus = 'active';
    if (d.statecode === 2) status = 'closed';
    else if (d.statecode === 3) status = 'paid';
    else if (d.statecode === 4) status = 'cancelled';

    return {
      id: d.invoiceid,
      name: d.name,
      invoiceNumber: d.invoicenumber,
      description: d.description,
      totalAmount: d.totalamount,
      totalLineItemAmount: d.totallineitemamount,
      totalDiscountAmount: d.totaldiscountamount,
      totalTax: d.totaltax,
      freightAmount: d.freightamount,
      discountPercentage: d.discountpercentage,
      discountAmount: d.discountamount,
      dueDate: d.duedate,
      dateDelivered: d.datedelivered,
      status,
      stateCode: d.statecode,
      isPriceLocked: d.ispricelocked,
      salesOrderId: d._salesorderid_value,
      opportunityId: d._opportunityid_value,
      customerId: d._customerid_value,
      priceLevelId: d._pricelevelid_value,
      ownerId: d._ownerid_value,
      createdAt: d.createdon,
      updatedAt: d.modifiedon,
    };
  }

  private readonly invoiceSelectFields =
    'invoiceid,name,invoicenumber,description,totalamount,totallineitemamount,totaldiscountamount,totaltax,freightamount,discountpercentage,discountamount,duedate,datedelivered,statuscode,statecode,ispricelocked,_salesorderid_value,_opportunityid_value,_customerid_value,_pricelevelid_value,_ownerid_value,createdon,modifiedon';

  async listInvoices(params?: PaginationParams): Promise<PaginatedResponse<Invoice>> {
    const limit = params?.limit || 20;
    const offset = params?.offset || 0;
    const queryParams = new URLSearchParams({
      $select: this.invoiceSelectFields,
      $top: String(limit),
      $skip: String(offset),
      $orderby: 'createdon desc',
    });
    let url = `/invoices?${queryParams}`;
    if (params?.cursor) url = params.cursor;
    const response = await this.request<DynamicsODataResponse<DynamicsInvoice>>(url);
    return {
      items: response.value.map((i) => this.mapInvoice(i)),
      count: response.value.length,
      total: response['@odata.count'],
      hasMore: !!response['@odata.nextLink'],
      nextCursor: response['@odata.nextLink'],
    };
  }

  async getInvoice(id: string): Promise<Invoice> {
    const invoice = await this.request<DynamicsInvoice>(`/invoices(${id})?$select=${this.invoiceSelectFields}`);
    return this.mapInvoice(invoice);
  }

  async createInvoice(input: InvoiceCreateInput): Promise<Invoice> {
    const body: Record<string, unknown> = { name: input.name, description: input.description };
    if (input.salesOrderId) body['salesorderid@odata.bind'] = `/salesorders(${input.salesOrderId})`;
    if (input.customerAccountId) body['customerid_account@odata.bind'] = `/accounts(${input.customerAccountId})`;
    if (input.priceLevelId) body['pricelevelid@odata.bind'] = `/pricelevels(${input.priceLevelId})`;
    if (input.dueDate) body.duedate = input.dueDate;
    if (input.customFields) Object.assign(body, input.customFields);
    const { id } = await this.createEntity('/invoices', body);
    return this.getInvoice(id);
  }

  async updateInvoice(id: string, input: Partial<InvoiceCreateInput>): Promise<Invoice> {
    const body: Record<string, unknown> = {};
    if (input.name !== undefined) body.name = input.name;
    if (input.description !== undefined) body.description = input.description;
    if (input.dueDate !== undefined) body.duedate = input.dueDate;
    if (input.customFields) Object.assign(body, input.customFields);
    await this.request(`/invoices(${id})`, { method: 'PATCH', body: JSON.stringify(body) });
    return this.getInvoice(id);
  }

  async deleteInvoice(id: string): Promise<void> {
    await this.request(`/invoices(${id})`, { method: 'DELETE' });
  }

  async lockInvoicePricing(id: string): Promise<void> {
    await this.request(`/invoices(${id})/Microsoft.Dynamics.CRM.LockInvoicePricing`, { method: 'POST', body: '{}' });
  }

  async cancelInvoice(id: string): Promise<void> {
    await this.request(`/invoices(${id})`, {
      method: 'PATCH',
      body: JSON.stringify({ statecode: 4, statuscode: 100003 }),
    });
  }

  // ===========================================================================
  // Products
  // ===========================================================================

  private mapProduct(d: DynamicsProduct): Product {
    let status: ProductStatus = 'draft';
    if (d.statecode === 0) status = 'active';
    else if (d.statecode === 1) status = 'retired';
    else if (d.statecode === 2) status = 'under_revision';

    return {
      id: d.productid,
      name: d.name,
      productNumber: d.productnumber,
      description: d.description,
      productStructure: d.productstructure,
      productTypeCode: d.producttypecode,
      quantityDecimal: d.quantitydecimal,
      currentCost: d.currentcost,
      standardCost: d.standardcost,
      price: d.price,
      stockWeight: d.stockweight,
      stockVolume: d.stockvolume,
      quantityOnHand: d.quantityonhand,
      defaultUomId: d._defaultuomid_value,
      defaultUomScheduleId: d._defaultuomscheduleid_value,
      subjectId: d._subjectid_value,
      status,
      stateCode: d.statecode,
      createdAt: d.createdon,
      updatedAt: d.modifiedon,
    };
  }

  private readonly productSelectFields =
    'productid,name,productnumber,description,productstructure,producttypecode,quantitydecimal,currentcost,standardcost,price,stockweight,stockvolume,quantityonhand,_defaultuomid_value,_defaultuomscheduleid_value,_subjectid_value,statuscode,statecode,createdon,modifiedon';

  async listProducts(params?: PaginationParams): Promise<PaginatedResponse<Product>> {
    const limit = params?.limit || 20;
    const offset = params?.offset || 0;
    const queryParams = new URLSearchParams({
      $select: this.productSelectFields,
      $top: String(limit),
      $skip: String(offset),
      $orderby: 'createdon desc',
    });
    let url = `/products?${queryParams}`;
    if (params?.cursor) url = params.cursor;
    const response = await this.request<DynamicsODataResponse<DynamicsProduct>>(url);
    return {
      items: response.value.map((p) => this.mapProduct(p)),
      count: response.value.length,
      total: response['@odata.count'],
      hasMore: !!response['@odata.nextLink'],
      nextCursor: response['@odata.nextLink'],
    };
  }

  async getProduct(id: string): Promise<Product> {
    const product = await this.request<DynamicsProduct>(`/products(${id})?$select=${this.productSelectFields}`);
    return this.mapProduct(product);
  }

  async createProduct(input: ProductCreateInput): Promise<Product> {
    const body: Record<string, unknown> = {
      name: input.name,
      productnumber: input.productNumber,
      description: input.description,
      productstructure: input.productStructure || 1,
      producttypecode: input.productTypeCode,
      quantitydecimal: input.quantityDecimal,
      currentcost: input.currentCost,
      standardcost: input.standardCost,
      price: input.price,
    };
    if (input.defaultUomId) body['defaultuomid@odata.bind'] = `/uoms(${input.defaultUomId})`;
    if (input.defaultUomScheduleId) body['defaultuomscheduleid@odata.bind'] = `/uomschedules(${input.defaultUomScheduleId})`;
    if (input.customFields) Object.assign(body, input.customFields);
    const { id } = await this.createEntity('/products', body);
    return this.getProduct(id);
  }

  async updateProduct(id: string, input: Partial<ProductCreateInput>): Promise<Product> {
    const body: Record<string, unknown> = {};
    if (input.name !== undefined) body.name = input.name;
    if (input.productNumber !== undefined) body.productnumber = input.productNumber;
    if (input.description !== undefined) body.description = input.description;
    if (input.productTypeCode !== undefined) body.producttypecode = input.productTypeCode;
    if (input.currentCost !== undefined) body.currentcost = input.currentCost;
    if (input.standardCost !== undefined) body.standardcost = input.standardCost;
    if (input.price !== undefined) body.price = input.price;
    if (input.customFields) Object.assign(body, input.customFields);
    await this.request(`/products(${id})`, { method: 'PATCH', body: JSON.stringify(body) });
    return this.getProduct(id);
  }

  async deleteProduct(id: string): Promise<void> {
    await this.request(`/products(${id})`, { method: 'DELETE' });
  }

  async publishProduct(id: string): Promise<void> {
    await this.request(`/products(${id})/Microsoft.Dynamics.CRM.PublishProductHierarchy`, { method: 'POST', body: '{}' });
  }

  // ===========================================================================
  // Price Lists
  // ===========================================================================

  private mapPriceLevel(d: DynamicsPriceLevel): PriceLevel {
    return {
      id: d.pricelevelid,
      name: d.name,
      description: d.description,
      beginDate: d.begindate,
      endDate: d.enddate,
      freightTermsCode: d.freighttermscode,
      paymentMethodCode: d.paymentmethodcode,
      shippingMethodCode: d.shippingmethodcode,
      status: d.statecode === 0 ? 'active' : 'inactive',
      stateCode: d.statecode,
      currencyId: d._transactioncurrencyid_value,
      createdAt: d.createdon,
      updatedAt: d.modifiedon,
    };
  }

  async listPriceLists(params?: PaginationParams): Promise<PaginatedResponse<PriceLevel>> {
    const limit = params?.limit || 20;
    const offset = params?.offset || 0;
    const queryParams = new URLSearchParams({
      $select: 'pricelevelid,name,description,begindate,enddate,freighttermscode,paymentmethodcode,shippingmethodcode,statuscode,statecode,_transactioncurrencyid_value,createdon,modifiedon',
      $top: String(limit),
      $skip: String(offset),
      $orderby: 'createdon desc',
    });
    let url = `/pricelevels?${queryParams}`;
    if (params?.cursor) url = params.cursor;
    const response = await this.request<DynamicsODataResponse<DynamicsPriceLevel>>(url);
    return {
      items: response.value.map((p) => this.mapPriceLevel(p)),
      count: response.value.length,
      total: response['@odata.count'],
      hasMore: !!response['@odata.nextLink'],
      nextCursor: response['@odata.nextLink'],
    };
  }

  async getPriceList(id: string): Promise<PriceLevel> {
    const priceLevel = await this.request<DynamicsPriceLevel>(`/pricelevels(${id})`);
    return this.mapPriceLevel(priceLevel);
  }

  async createPriceList(input: { name: string; description?: string; currencyId?: string }): Promise<PriceLevel> {
    const body: Record<string, unknown> = { name: input.name, description: input.description };
    if (input.currencyId) body['transactioncurrencyid@odata.bind'] = `/transactioncurrencies(${input.currencyId})`;
    const { id } = await this.createEntity('/pricelevels', body);
    return this.getPriceList(id);
  }

  // ===========================================================================
  // Competitors
  // ===========================================================================

  private mapCompetitor(d: DynamicsCompetitor): Competitor {
    return {
      id: d.competitorid,
      name: d.name,
      website: d.websiteurl,
      tickerSymbol: d.tickersymbol,
      stockExchange: d.stockexchange,
      reportedRevenue: d.reportedrevenue,
      reportingQuarter: d.reportingquarter,
      reportingYear: d.reportingyear,
      keyProduct: d.keyproduct,
      strengths: d.strengths,
      weaknesses: d.weaknesses,
      overview: d.overview,
      opportunities: d.opportunities,
      threats: d.threats,
      winPercentage: d.winpercentage,
      address: {
        street: d.address1_line1,
        city: d.address1_city,
        state: d.address1_stateorprovince,
        postalCode: d.address1_postalcode,
        country: d.address1_country,
      },
      status: d.statecode === 0 ? 'active' : 'inactive',
      stateCode: d.statecode,
      createdAt: d.createdon,
      updatedAt: d.modifiedon,
    };
  }

  private readonly competitorSelectFields =
    'competitorid,name,websiteurl,tickersymbol,stockexchange,reportedrevenue,reportingquarter,reportingyear,keyproduct,strengths,weaknesses,overview,opportunities,threats,winpercentage,address1_line1,address1_city,address1_stateorprovince,address1_postalcode,address1_country,statuscode,statecode,createdon,modifiedon';

  async listCompetitors(params?: PaginationParams): Promise<PaginatedResponse<Competitor>> {
    const limit = params?.limit || 20;
    const offset = params?.offset || 0;
    const queryParams = new URLSearchParams({
      $select: this.competitorSelectFields,
      $top: String(limit),
      $skip: String(offset),
      $orderby: 'createdon desc',
    });
    let url = `/competitors?${queryParams}`;
    if (params?.cursor) url = params.cursor;
    const response = await this.request<DynamicsODataResponse<DynamicsCompetitor>>(url);
    return {
      items: response.value.map((c) => this.mapCompetitor(c)),
      count: response.value.length,
      total: response['@odata.count'],
      hasMore: !!response['@odata.nextLink'],
      nextCursor: response['@odata.nextLink'],
    };
  }

  async getCompetitor(id: string): Promise<Competitor> {
    const competitor = await this.request<DynamicsCompetitor>(`/competitors(${id})?$select=${this.competitorSelectFields}`);
    return this.mapCompetitor(competitor);
  }

  async createCompetitor(input: CompetitorCreateInput): Promise<Competitor> {
    const body: Record<string, unknown> = {
      name: input.name,
      websiteurl: input.website,
      tickersymbol: input.tickerSymbol,
      keyproduct: input.keyProduct,
      strengths: input.strengths,
      weaknesses: input.weaknesses,
      overview: input.overview,
    };
    if (input.customFields) Object.assign(body, input.customFields);
    const { id } = await this.createEntity('/competitors', body);
    return this.getCompetitor(id);
  }

  async updateCompetitor(id: string, input: Partial<CompetitorCreateInput>): Promise<Competitor> {
    const body: Record<string, unknown> = {};
    if (input.name !== undefined) body.name = input.name;
    if (input.website !== undefined) body.websiteurl = input.website;
    if (input.tickerSymbol !== undefined) body.tickersymbol = input.tickerSymbol;
    if (input.keyProduct !== undefined) body.keyproduct = input.keyProduct;
    if (input.strengths !== undefined) body.strengths = input.strengths;
    if (input.weaknesses !== undefined) body.weaknesses = input.weaknesses;
    if (input.overview !== undefined) body.overview = input.overview;
    if (input.customFields) Object.assign(body, input.customFields);
    await this.request(`/competitors(${id})`, { method: 'PATCH', body: JSON.stringify(body) });
    return this.getCompetitor(id);
  }

  async deleteCompetitor(id: string): Promise<void> {
    await this.request(`/competitors(${id})`, { method: 'DELETE' });
  }

  async associateCompetitorToOpportunity(competitorId: string, opportunityId: string): Promise<void> {
    await this.request(`/opportunities(${opportunityId})/opportunitycompetitors_association/$ref`, {
      method: 'POST',
      body: JSON.stringify({ '@odata.id': `${this.baseUrl}/competitors(${competitorId})` }),
    });
  }

  async disassociateCompetitorFromOpportunity(competitorId: string, opportunityId: string): Promise<void> {
    await this.request(`/opportunities(${opportunityId})/opportunitycompetitors_association(${competitorId})/$ref`, {
      method: 'DELETE',
    });
  }

  async listOpportunityCompetitors(opportunityId: string): Promise<Competitor[]> {
    const response = await this.request<DynamicsODataResponse<DynamicsCompetitor>>(
      `/opportunities(${opportunityId})/opportunitycompetitors_association?$select=${this.competitorSelectFields}`
    );
    return response.value.map((c) => this.mapCompetitor(c));
  }

  // ===========================================================================
  // Campaigns
  // ===========================================================================

  private mapCampaign(d: DynamicsCampaign): Campaign {
    return {
      id: d.campaignid,
      name: d.name,
      codeName: d.codename,
      description: d.description,
      message: d.message,
      objective: d.objective,
      typeCode: d.typecode,
      proposedStart: d.proposedstart,
      proposedEnd: d.proposedend,
      actualStart: d.actualstart,
      actualEnd: d.actualend,
      budgetedCost: d.budgetedcost,
      otherCost: d.othercost,
      totalCampaignActivityActualCost: d.totalcampaignactivityactualcost,
      totalActualCost: d.totalactualcost,
      expectedResponse: d.expectedresponse,
      expectedRevenue: d.expectedrevenue,
      status: d.statecode === 0 ? 'active' : 'inactive',
      stateCode: d.statecode,
      priceLevelId: d._pricelevelid_value,
      ownerId: d._ownerid_value,
      currencyId: d._transactioncurrencyid_value,
      createdAt: d.createdon,
      updatedAt: d.modifiedon,
    };
  }

  private readonly campaignSelectFields =
    'campaignid,name,codename,description,message,objective,typecode,proposedstart,proposedend,actualstart,actualend,budgetedcost,othercost,totalcampaignactivityactualcost,totalactualcost,expectedresponse,expectedrevenue,statuscode,statecode,_pricelevelid_value,_ownerid_value,_transactioncurrencyid_value,createdon,modifiedon';

  async listCampaigns(params?: PaginationParams): Promise<PaginatedResponse<Campaign>> {
    const limit = params?.limit || 20;
    const offset = params?.offset || 0;
    const queryParams = new URLSearchParams({
      $select: this.campaignSelectFields,
      $top: String(limit),
      $skip: String(offset),
      $orderby: 'createdon desc',
    });
    let url = `/campaigns?${queryParams}`;
    if (params?.cursor) url = params.cursor;
    const response = await this.request<DynamicsODataResponse<DynamicsCampaign>>(url);
    return {
      items: response.value.map((c) => this.mapCampaign(c)),
      count: response.value.length,
      total: response['@odata.count'],
      hasMore: !!response['@odata.nextLink'],
      nextCursor: response['@odata.nextLink'],
    };
  }

  async getCampaign(id: string): Promise<Campaign> {
    const campaign = await this.request<DynamicsCampaign>(`/campaigns(${id})?$select=${this.campaignSelectFields}`);
    return this.mapCampaign(campaign);
  }

  async createCampaign(input: CampaignCreateInput): Promise<Campaign> {
    const body: Record<string, unknown> = {
      name: input.name,
      codename: input.codeName,
      description: input.description,
      message: input.message,
      objective: input.objective,
      typecode: input.typeCode,
      proposedstart: input.proposedStart,
      proposedend: input.proposedEnd,
      budgetedcost: input.budgetedCost,
    };
    if (input.priceLevelId) body['pricelevelid@odata.bind'] = `/pricelevels(${input.priceLevelId})`;
    if (input.currencyId) body['transactioncurrencyid@odata.bind'] = `/transactioncurrencies(${input.currencyId})`;
    if (input.customFields) Object.assign(body, input.customFields);
    const { id } = await this.createEntity('/campaigns', body);
    return this.getCampaign(id);
  }

  async updateCampaign(id: string, input: Partial<CampaignCreateInput>): Promise<Campaign> {
    const body: Record<string, unknown> = {};
    if (input.name !== undefined) body.name = input.name;
    if (input.codeName !== undefined) body.codename = input.codeName;
    if (input.description !== undefined) body.description = input.description;
    if (input.message !== undefined) body.message = input.message;
    if (input.objective !== undefined) body.objective = input.objective;
    if (input.typeCode !== undefined) body.typecode = input.typeCode;
    if (input.proposedStart !== undefined) body.proposedstart = input.proposedStart;
    if (input.proposedEnd !== undefined) body.proposedend = input.proposedEnd;
    if (input.budgetedCost !== undefined) body.budgetedcost = input.budgetedCost;
    if (input.customFields) Object.assign(body, input.customFields);
    await this.request(`/campaigns(${id})`, { method: 'PATCH', body: JSON.stringify(body) });
    return this.getCampaign(id);
  }

  async deleteCampaign(id: string): Promise<void> {
    await this.request(`/campaigns(${id})`, { method: 'DELETE' });
  }

  async listCampaignActivities(campaignId: string): Promise<unknown[]> {
    const response = await this.request<DynamicsODataResponse<unknown>>(
      `/campaigns(${campaignId})/Campaign_CampaignActivities?$select=activityid,subject,description,channeltypecode,typecode,scheduledstart,scheduledend,budgetedcost,actualcost,statuscode,statecode,createdon`
    );
    return response.value;
  }

  async createCampaignActivity(input: {
    campaignId: string;
    subject: string;
    description?: string;
    channelTypeCode?: number;
    typeCode?: number;
    scheduledStart?: string;
    scheduledEnd?: string;
    budgetedCost?: number;
  }): Promise<unknown> {
    const body: Record<string, unknown> = {
      subject: input.subject,
      description: input.description,
      channeltypecode: input.channelTypeCode,
      typecode: input.typeCode,
      scheduledstart: input.scheduledStart,
      scheduledend: input.scheduledEnd,
      budgetedcost: input.budgetedCost,
      'regardingobjectid_campaign@odata.bind': `/campaigns(${input.campaignId})`,
    };
    const { id } = await this.createEntity('/campaignactivities', body);
    return { id };
  }

  async listCampaignResponses(campaignId: string): Promise<unknown[]> {
    const response = await this.request<DynamicsODataResponse<unknown>>(
      `/campaigns(${campaignId})/Campaign_CampaignResponses?$select=activityid,subject,description,channeltypecode,responsecode,receivedon,firstname,lastname,emailaddress,telephone,companyname,statuscode,statecode,createdon`
    );
    return response.value;
  }

  async createCampaignResponse(input: {
    campaignId: string;
    subject: string;
    description?: string;
    channelTypeCode?: number;
    responseCode?: number;
    receivedOn?: string;
    firstName?: string;
    lastName?: string;
    emailAddress?: string;
    telephone?: string;
    companyName?: string;
  }): Promise<unknown> {
    const body: Record<string, unknown> = {
      subject: input.subject,
      description: input.description,
      channeltypecode: input.channelTypeCode,
      responsecode: input.responseCode,
      receivedon: input.receivedOn,
      firstname: input.firstName,
      lastname: input.lastName,
      emailaddress: input.emailAddress,
      telephone: input.telephone,
      companyname: input.companyName,
      'regardingobjectid_campaign@odata.bind': `/campaigns(${input.campaignId})`,
    };
    const { id } = await this.createEntity('/campaignresponses', body);
    return { id };
  }

  async addCampaignMember(campaignId: string, entityType: string, entityId: string): Promise<void> {
    // Use AddMemberList action or associate directly based on entity type
    const entitySetName = entityType === 'contact' ? 'contacts' : entityType === 'lead' ? 'leads' : 'accounts';
    await this.request(`/campaigns(${campaignId})/Campaign_${entityType}s/$ref`, {
      method: 'POST',
      body: JSON.stringify({ '@odata.id': `${this.baseUrl}/${entitySetName}(${entityId})` }),
    });
  }

  async removeCampaignMember(campaignId: string, entityType: string, entityId: string): Promise<void> {
    await this.request(`/campaigns(${campaignId})/Campaign_${entityType}s(${entityId})/$ref`, {
      method: 'DELETE',
    });
  }

  // ===========================================================================
  // Cases (Incidents)
  // ===========================================================================

  private mapCase(d: DynamicsIncident): Case {
    let status: CaseStatus = 'active';
    if (d.statecode === 1) status = 'resolved';
    else if (d.statecode === 2) status = 'cancelled';

    return {
      id: d.incidentid,
      title: d.title,
      ticketNumber: d.ticketnumber,
      description: d.description,
      caseOriginCode: d.caseorigincode,
      caseTypeCode: d.casetypecode,
      priorityCode: d.prioritycode,
      severityCode: d.severitycode,
      status,
      stateCode: d.statecode,
      escalatedOn: d.escalatedon,
      isEscalated: d.isescalated,
      followUpBy: d.followupby,
      customerId: d._customerid_value,
      primaryContactId: d._primarycontactid_value,
      productId: d._productid_value,
      subjectId: d._subjectid_value,
      ownerId: d._ownerid_value,
      entitlementId: d._entitlementid_value,
      contractId: d._contractid_value,
      contractDetailId: d._contractdetailid_value,
      resolveBy: d.resolveby,
      responseBy: d.responseby,
      createdAt: d.createdon,
      updatedAt: d.modifiedon,
    };
  }

  private readonly caseSelectFields =
    'incidentid,title,ticketnumber,description,caseorigincode,casetypecode,prioritycode,severitycode,statuscode,statecode,escalatedon,isescalated,followupby,_customerid_value,_primarycontactid_value,_productid_value,_subjectid_value,_ownerid_value,_entitlementid_value,_contractid_value,_contractdetailid_value,resolveby,responseby,createdon,modifiedon';

  async listCases(params?: PaginationParams): Promise<PaginatedResponse<Case>> {
    const limit = params?.limit || 20;
    const offset = params?.offset || 0;
    const queryParams = new URLSearchParams({
      $select: this.caseSelectFields,
      $top: String(limit),
      $skip: String(offset),
      $orderby: 'createdon desc',
    });
    let url = `/incidents?${queryParams}`;
    if (params?.cursor) url = params.cursor;
    const response = await this.request<DynamicsODataResponse<DynamicsIncident>>(url);
    return {
      items: response.value.map((c) => this.mapCase(c)),
      count: response.value.length,
      total: response['@odata.count'],
      hasMore: !!response['@odata.nextLink'],
      nextCursor: response['@odata.nextLink'],
    };
  }

  async getCase(id: string): Promise<Case> {
    const incident = await this.request<DynamicsIncident>(`/incidents(${id})?$select=${this.caseSelectFields}`);
    return this.mapCase(incident);
  }

  async createCase(input: CaseCreateInput): Promise<Case> {
    const body: Record<string, unknown> = {
      title: input.title,
      description: input.description,
      caseorigincode: input.caseOriginCode,
      casetypecode: input.caseTypeCode,
      prioritycode: input.priorityCode,
      severitycode: input.severityCode,
    };
    if (input.customerAccountId) body['customerid_account@odata.bind'] = `/accounts(${input.customerAccountId})`;
    if (input.customerContactId) body['customerid_contact@odata.bind'] = `/contacts(${input.customerContactId})`;
    if (input.primaryContactId) body['primarycontactid@odata.bind'] = `/contacts(${input.primaryContactId})`;
    if (input.productId) body['productid@odata.bind'] = `/products(${input.productId})`;
    if (input.subjectId) body['subjectid@odata.bind'] = `/subjects(${input.subjectId})`;
    if (input.customFields) Object.assign(body, input.customFields);
    const { id } = await this.createEntity('/incidents', body);
    return this.getCase(id);
  }

  async updateCase(id: string, input: Partial<CaseCreateInput>): Promise<Case> {
    const body: Record<string, unknown> = {};
    if (input.title !== undefined) body.title = input.title;
    if (input.description !== undefined) body.description = input.description;
    if (input.caseOriginCode !== undefined) body.caseorigincode = input.caseOriginCode;
    if (input.caseTypeCode !== undefined) body.casetypecode = input.caseTypeCode;
    if (input.priorityCode !== undefined) body.prioritycode = input.priorityCode;
    if (input.severityCode !== undefined) body.severitycode = input.severityCode;
    if (input.customFields) Object.assign(body, input.customFields);
    await this.request(`/incidents(${id})`, { method: 'PATCH', body: JSON.stringify(body) });
    return this.getCase(id);
  }

  async deleteCase(id: string): Promise<void> {
    await this.request(`/incidents(${id})`, { method: 'DELETE' });
  }

  async resolveCase(id: string, resolution: { subject: string; description?: string; timeSpent?: number }): Promise<void> {
    await this.request(`/incidents(${id})/Microsoft.Dynamics.CRM.CloseIncident`, {
      method: 'POST',
      body: JSON.stringify({
        IncidentResolution: {
          subject: resolution.subject,
          description: resolution.description,
          timespent: resolution.timeSpent,
          'incidentid@odata.bind': `/incidents(${id})`,
        },
        Status: -1,
      }),
    });
  }

  async cancelCase(id: string): Promise<void> {
    await this.request(`/incidents(${id})`, {
      method: 'PATCH',
      body: JSON.stringify({ statecode: 2, statuscode: 6 }),
    });
  }

  async reactivateCase(id: string): Promise<void> {
    await this.request(`/incidents(${id})`, {
      method: 'PATCH',
      body: JSON.stringify({ statecode: 0, statuscode: 1 }), // Active, In Progress
    });
  }

  // ===========================================================================
  // Goals
  // ===========================================================================

  private mapGoal(d: DynamicsGoal): Goal {
    return {
      id: d.goalid,
      title: d.title,
      goalOwnerId: d._goalownerid_value,
      metricId: d._metricid_value,
      targetMoney: d.targetmoney,
      targetDecimal: d.targetdecimal,
      targetInteger: d.targetinteger,
      actualMoney: d.actualmoney,
      actualDecimal: d.actualdecimal,
      actualInteger: d.actualinteger,
      inProgressMoney: d.inprogressmoney,
      inProgressDecimal: d.inprogressdecimal,
      inProgressInteger: d.inprogressinteger,
      percentage: d.percentage,
      fiscalPeriod: d.fiscalperiod,
      fiscalYear: d.fiscalyear,
      goalStartDate: d.goalstartdate,
      goalEndDate: d.goalenddate,
      considerOnlyGoalOwnersRecords: d.consideronlygoalownersrecords,
      parentGoalId: d._parentgoalid_value,
      status: d.statecode === 0 ? 'active' : 'inactive',
      stateCode: d.statecode,
      lastRolledUpDate: d.lastrolledupdate,
      createdAt: d.createdon,
      updatedAt: d.modifiedon,
    };
  }

  private readonly goalSelectFields =
    'goalid,title,_goalownerid_value,_metricid_value,targetmoney,targetdecimal,targetinteger,actualmoney,actualdecimal,actualinteger,inprogressmoney,inprogressdecimal,inprogressinteger,percentage,fiscalperiod,fiscalyear,goalstartdate,goalenddate,consideronlygoalownersrecords,_parentgoalid_value,statuscode,statecode,lastrolledupdate,createdon,modifiedon';

  async listGoals(params?: PaginationParams): Promise<PaginatedResponse<Goal>> {
    const limit = params?.limit || 20;
    const offset = params?.offset || 0;
    const queryParams = new URLSearchParams({
      $select: this.goalSelectFields,
      $top: String(limit),
      $skip: String(offset),
      $orderby: 'createdon desc',
    });
    let url = `/goals?${queryParams}`;
    if (params?.cursor) url = params.cursor;
    const response = await this.request<DynamicsODataResponse<DynamicsGoal>>(url);
    return {
      items: response.value.map((g) => this.mapGoal(g)),
      count: response.value.length,
      total: response['@odata.count'],
      hasMore: !!response['@odata.nextLink'],
      nextCursor: response['@odata.nextLink'],
    };
  }

  async getGoal(id: string): Promise<Goal> {
    const goal = await this.request<DynamicsGoal>(`/goals(${id})?$select=${this.goalSelectFields}`);
    return this.mapGoal(goal);
  }

  async createGoal(input: GoalCreateInput): Promise<Goal> {
    const body: Record<string, unknown> = {
      title: input.title,
      'goalownerid@odata.bind': `/systemusers(${input.goalOwnerId})`,
      'metricid@odata.bind': `/metrics(${input.metricId})`,
      targetmoney: input.targetMoney,
      targetdecimal: input.targetDecimal,
      targetinteger: input.targetInteger,
      goalstartdate: input.goalStartDate,
      goalenddate: input.goalEndDate,
      fiscalperiod: input.fiscalPeriod,
      fiscalyear: input.fiscalYear,
      consideronlygoalownersrecords: input.considerOnlyGoalOwnersRecords,
    };
    if (input.parentGoalId) body['parentgoalid@odata.bind'] = `/goals(${input.parentGoalId})`;
    if (input.customFields) Object.assign(body, input.customFields);
    const { id } = await this.createEntity('/goals', body);
    return this.getGoal(id);
  }

  async updateGoal(id: string, input: Partial<GoalCreateInput>): Promise<Goal> {
    const body: Record<string, unknown> = {};
    if (input.title !== undefined) body.title = input.title;
    if (input.targetMoney !== undefined) body.targetmoney = input.targetMoney;
    if (input.targetDecimal !== undefined) body.targetdecimal = input.targetDecimal;
    if (input.targetInteger !== undefined) body.targetinteger = input.targetInteger;
    if (input.goalStartDate !== undefined) body.goalstartdate = input.goalStartDate;
    if (input.goalEndDate !== undefined) body.goalenddate = input.goalEndDate;
    if (input.fiscalPeriod !== undefined) body.fiscalperiod = input.fiscalPeriod;
    if (input.fiscalYear !== undefined) body.fiscalyear = input.fiscalYear;
    if (input.considerOnlyGoalOwnersRecords !== undefined) body.consideronlygoalownersrecords = input.considerOnlyGoalOwnersRecords;
    if (input.customFields) Object.assign(body, input.customFields);
    await this.request(`/goals(${id})`, { method: 'PATCH', body: JSON.stringify(body) });
    return this.getGoal(id);
  }

  async deleteGoal(id: string): Promise<void> {
    await this.request(`/goals(${id})`, { method: 'DELETE' });
  }

  async recalculateGoal(id: string): Promise<void> {
    await this.request(`/goals(${id})/Microsoft.Dynamics.CRM.Recalculate`, { method: 'POST', body: '{}' });
  }

  async listGoalMetrics(params?: PaginationParams): Promise<PaginatedResponse<GoalMetric>> {
    const limit = params?.limit || 20;
    const offset = params?.offset || 0;
    const queryParams = new URLSearchParams({
      $select: 'metricid,name,description,amountdatatype,isamount,isstretchtracked,statuscode,statecode',
      $top: String(limit),
      $skip: String(offset),
    });
    let url = `/metrics?${queryParams}`;
    if (params?.cursor) url = params.cursor;
    const response = await this.request<DynamicsODataResponse<{ metricid: string; name: string; description?: string; amountdatatype?: number; isamount?: boolean; isstretchtracked?: boolean; statuscode?: number; statecode?: number }>>(url);
    return {
      items: response.value.map((m) => ({
        id: m.metricid,
        name: m.name,
        description: m.description,
        amountDataType: m.amountdatatype,
        isAmount: m.isamount,
        isStretchTracked: m.isstretchtracked,
        status: m.statecode === 0 ? 'active' : 'inactive',
        stateCode: m.statecode,
      })),
      count: response.value.length,
      total: response['@odata.count'],
      hasMore: !!response['@odata.nextLink'],
      nextCursor: response['@odata.nextLink'],
    };
  }

  async getGoalMetric(id: string): Promise<GoalMetric> {
    const m = await this.request<{ metricid: string; name: string; description?: string; amountdatatype?: number; isamount?: boolean; isstretchtracked?: boolean; statuscode?: number; statecode?: number }>(
      `/metrics(${id})?$select=metricid,name,description,amountdatatype,isamount,isstretchtracked,statuscode,statecode`
    );
    return {
      id: m.metricid,
      name: m.name,
      description: m.description,
      amountDataType: m.amountdatatype,
      isAmount: m.isamount,
      isStretchTracked: m.isstretchtracked,
      status: m.statecode === 0 ? 'active' : 'inactive',
      stateCode: m.statecode,
    };
  }

  // ===========================================================================
  // Notes (Annotations)
  // ===========================================================================

  private mapNote(d: DynamicsAnnotation): Note {
    return {
      id: d.annotationid,
      subject: d.subject,
      noteText: d.notetext,
      objectId: d._objectid_value,
      objectTypeCode: d.objecttypecode,
      isDocument: d.isdocument,
      fileName: d.filename,
      mimeType: d.mimetype,
      fileSize: d.filesize,
      documentBody: d.documentbody,
      ownerId: d._ownerid_value,
      createdById: d._createdby_value,
      createdAt: d.createdon,
      updatedAt: d.modifiedon,
    };
  }

  private readonly noteSelectFields =
    'annotationid,subject,notetext,_objectid_value,objecttypecode,isdocument,filename,mimetype,filesize,_ownerid_value,_createdby_value,createdon,modifiedon';

  async listNotes(params?: PaginationParams & { regardingEntityType?: string; regardingId?: string }): Promise<PaginatedResponse<Note>> {
    const limit = params?.limit || 20;
    const offset = params?.offset || 0;
    const queryParams = new URLSearchParams({
      $select: this.noteSelectFields,
      $top: String(limit),
      $skip: String(offset),
      $orderby: 'createdon desc',
    });
    if (params?.regardingId) {
      queryParams.set('$filter', `_objectid_value eq ${params.regardingId}`);
    }
    let url = `/annotations?${queryParams}`;
    if (params?.cursor) url = params.cursor;
    const response = await this.request<DynamicsODataResponse<DynamicsAnnotation>>(url);
    return {
      items: response.value.map((n) => this.mapNote(n)),
      count: response.value.length,
      total: response['@odata.count'],
      hasMore: !!response['@odata.nextLink'],
      nextCursor: response['@odata.nextLink'],
    };
  }

  async getNote(id: string): Promise<Note> {
    const note = await this.request<DynamicsAnnotation>(`/annotations(${id})?$select=${this.noteSelectFields}`);
    return this.mapNote(note);
  }

  async createNote(input: NoteCreateInput): Promise<Note> {
    const body: Record<string, unknown> = {
      subject: input.subject,
      notetext: input.noteText,
      isdocument: !!(input.fileName || input.documentBody),
      filename: input.fileName,
      mimetype: input.mimeType,
      documentbody: input.documentBody,
    };
    // Set objectid based on object type
    if (input.regardingEntityType && input.regardingId) {
      body[`objectid_${input.regardingEntityType}@odata.bind`] = `/${input.regardingEntityType}s(${input.regardingId})`;
    } else if (input.objectType && input.objectId) {
      body[`objectid_${input.objectType}@odata.bind`] = `/${input.objectType}s(${input.objectId})`;
    }
    if (input.customFields) Object.assign(body, input.customFields);
    const { id } = await this.createEntity('/annotations', body);
    return this.getNote(id);
  }

  async updateNote(id: string, input: { subject?: string; noteText?: string; fileName?: string; mimeType?: string; documentBody?: string }): Promise<Note> {
    const body: Record<string, unknown> = {};
    if (input.subject !== undefined) body.subject = input.subject;
    if (input.noteText !== undefined) body.notetext = input.noteText;
    if (input.fileName !== undefined) body.filename = input.fileName;
    if (input.mimeType !== undefined) body.mimetype = input.mimeType;
    if (input.documentBody !== undefined) {
      body.documentbody = input.documentBody;
      body.isdocument = true;
    }
    await this.request(`/annotations(${id})`, { method: 'PATCH', body: JSON.stringify(body) });
    return this.getNote(id);
  }

  async deleteNote(id: string): Promise<void> {
    await this.request(`/annotations(${id})`, { method: 'DELETE' });
  }

  async getNoteAttachment(id: string): Promise<{ fileName: string; mimeType: string; documentBody: string; fileSize: number }> {
    const note = await this.request<DynamicsAnnotation>(`/annotations(${id})?$select=filename,mimetype,documentbody,filesize`);
    return {
      fileName: note.filename || '',
      mimeType: note.mimetype || '',
      documentBody: note.documentbody || '',
      fileSize: note.filesize || 0,
    };
  }

  async listEntityNotes(entityType: string, entityId: string, limit = 20): Promise<Note[]> {
    const response = await this.request<DynamicsODataResponse<DynamicsAnnotation>>(
      `/annotations?$filter=_objectid_value eq ${entityId}&$select=${this.noteSelectFields}&$top=${limit}&$orderby=createdon desc`
    );
    return response.value.map((n) => this.mapNote(n));
  }

  async addAttachmentToNote(noteId: string, input: { fileName: string; mimeType: string; documentBody: string }): Promise<void> {
    await this.request(`/annotations(${noteId})`, {
      method: 'PATCH',
      body: JSON.stringify({
        isdocument: true,
        filename: input.fileName,
        mimetype: input.mimeType,
        documentbody: input.documentBody,
      }),
    });
  }

  async removeAttachmentFromNote(noteId: string): Promise<void> {
    await this.request(`/annotations(${noteId})`, {
      method: 'PATCH',
      body: JSON.stringify({
        isdocument: false,
        filename: null,
        mimetype: null,
        documentbody: null,
      }),
    });
  }

  async searchNotes(query: string, limit = 20): Promise<Note[]> {
    const response = await this.request<DynamicsODataResponse<DynamicsAnnotation>>(
      `/annotations?$filter=contains(subject,'${query}') or contains(notetext,'${query}')&$select=${this.noteSelectFields}&$top=${limit}&$orderby=createdon desc`
    );
    return response.value.map((n) => this.mapNote(n));
  }

  // ===========================================================================
  // Users
  // ===========================================================================

  private mapUser(d: DynamicsSystemUser): SystemUser {
    return {
      id: d.systemuserid,
      fullName: d.fullname,
      firstName: d.firstname,
      lastName: d.lastname,
      domainName: d.domainname,
      internalEmailAddress: d.internalemailaddress,
      title: d.title,
      jobTitle: d.jobtitle,
      mobilePhone: d.mobilephone,
      phone: d.address1_telephone1,
      businessUnitId: d._businessunitid_value,
      territoryId: d._territoryid_value,
      positionId: d._positionid_value,
      queueId: d._queueid_value,
      isDisabled: d.isdisabled,
      accessMode: d.accessmode,
      userLicenseType: d.userlicensetype,
      setupUser: d.setupuser,
      createdAt: d.createdon,
      updatedAt: d.modifiedon,
    };
  }

  private readonly userSelectFields =
    'systemuserid,fullname,firstname,lastname,domainname,internalemailaddress,title,jobtitle,mobilephone,address1_telephone1,_businessunitid_value,_territoryid_value,_positionid_value,_queueid_value,isdisabled,accessmode,userlicensetype,setupuser,createdon,modifiedon';

  async listUsers(params?: PaginationParams & { activeOnly?: boolean }): Promise<PaginatedResponse<SystemUser>> {
    const limit = params?.limit || 20;
    const offset = params?.offset || 0;
    const activeOnly = params?.activeOnly !== false; // Default to true
    const filters: string[] = [];
    if (activeOnly) {
      filters.push('isdisabled eq false');
    }
    const queryParams = new URLSearchParams({
      $select: this.userSelectFields,
      $top: String(limit),
      $skip: String(offset),
      $orderby: 'fullname asc',
    });
    if (filters.length > 0) {
      queryParams.set('$filter', filters.join(' and '));
    }
    let url = `/systemusers?${queryParams}`;
    if (params?.cursor) url = params.cursor;
    const response = await this.request<DynamicsODataResponse<DynamicsSystemUser>>(url);
    return {
      items: response.value.map((u) => this.mapUser(u)),
      count: response.value.length,
      total: response['@odata.count'],
      hasMore: !!response['@odata.nextLink'],
      nextCursor: response['@odata.nextLink'],
    };
  }

  async getUser(id: string): Promise<SystemUser> {
    const user = await this.request<DynamicsSystemUser>(`/systemusers(${id})?$select=${this.userSelectFields}`);
    return this.mapUser(user);
  }

  async getCurrentUser(): Promise<SystemUser> {
    const whoAmI = await this.request<DynamicsWhoAmI>('/WhoAmI');
    return this.getUser(whoAmI.UserId);
  }

  async searchUsers(query: string, limit = 20): Promise<SystemUser[]> {
    const filters: string[] = ['isdisabled eq false'];
    if (query) {
      filters.push(`(contains(fullname,'${query}') or contains(internalemailaddress,'${query}'))`);
    }
    const queryParams = new URLSearchParams({
      $select: this.userSelectFields,
      $top: String(limit),
      $filter: filters.join(' and '),
      $orderby: 'fullname asc',
    });
    const response = await this.request<DynamicsODataResponse<DynamicsSystemUser>>(`/systemusers?${queryParams}`);
    return response.value.map((u) => this.mapUser(u));
  }

  // ===========================================================================
  // Teams
  // ===========================================================================

  private mapTeam(d: DynamicsTeam): Team {
    return {
      id: d.teamid,
      name: d.name,
      description: d.description,
      teamType: d.teamtype,
      businessUnitId: d._businessunitid_value,
      administratorId: d._administratorid_value,
      queueId: d._queueid_value,
      isDefault: d.isdefault,
      createdAt: d.createdon,
    };
  }

  async listTeams(params?: PaginationParams): Promise<PaginatedResponse<Team>> {
    const limit = params?.limit || 20;
    const offset = params?.offset || 0;
    const queryParams = new URLSearchParams({
      $select: 'teamid,name,description,teamtype,_businessunitid_value,_administratorid_value,_queueid_value,isdefault,createdon',
      $top: String(limit),
      $skip: String(offset),
      $orderby: 'name asc',
    });
    let url = `/teams?${queryParams}`;
    if (params?.cursor) url = params.cursor;
    const response = await this.request<DynamicsODataResponse<DynamicsTeam>>(url);
    return {
      items: response.value.map((t) => this.mapTeam(t)),
      count: response.value.length,
      total: response['@odata.count'],
      hasMore: !!response['@odata.nextLink'],
      nextCursor: response['@odata.nextLink'],
    };
  }

  async getTeam(id: string): Promise<Team> {
    const team = await this.request<DynamicsTeam>(`/teams(${id})`);
    return this.mapTeam(team);
  }

  async createTeam(input: { name: string; description?: string; businessUnitId: string; teamType?: number; administratorId?: string }): Promise<Team> {
    const body: Record<string, unknown> = {
      name: input.name,
      description: input.description,
      teamtype: input.teamType || 0, // 0 = Owner team
      'businessunitid@odata.bind': `/businessunits(${input.businessUnitId})`,
    };
    if (input.administratorId) {
      body['administratorid@odata.bind'] = `/systemusers(${input.administratorId})`;
    }
    const { id } = await this.createEntity('/teams', body);
    return this.getTeam(id);
  }

  async updateTeam(id: string, input: { name?: string; description?: string; administratorId?: string }): Promise<Team> {
    const body: Record<string, unknown> = {};
    if (input.name !== undefined) body.name = input.name;
    if (input.description !== undefined) body.description = input.description;
    if (input.administratorId) {
      body['administratorid@odata.bind'] = `/systemusers(${input.administratorId})`;
    }
    await this.request(`/teams(${id})`, {
      method: 'PATCH',
      body: JSON.stringify(body),
    });
    return this.getTeam(id);
  }

  async deleteTeam(id: string): Promise<void> {
    await this.request(`/teams(${id})`, { method: 'DELETE' });
  }

  async addTeamMember(teamId: string, userId: string): Promise<void> {
    await this.request(`/teams(${teamId})/teammembership_association/$ref`, {
      method: 'POST',
      body: JSON.stringify({ '@odata.id': `${this.baseUrl}/systemusers(${userId})` }),
    });
  }

  async removeTeamMember(teamId: string, userId: string): Promise<void> {
    await this.request(`/teams(${teamId})/teammembership_association(${userId})/$ref`, {
      method: 'DELETE',
    });
  }

  async listTeamMembers(teamId: string): Promise<SystemUser[]> {
    const response = await this.request<DynamicsODataResponse<DynamicsSystemUser>>(
      `/teams(${teamId})/teammembership_association?$select=${this.userSelectFields}`
    );
    return response.value.map((u) => this.mapUser(u));
  }

  // ===========================================================================
  // Business Units
  // ===========================================================================

  private mapBusinessUnit(d: DynamicsBusinessUnit): BusinessUnit {
    return {
      id: d.businessunitid,
      name: d.name,
      parentBusinessUnitId: d._parentbusinessunitid_value,
      divisionName: d.divisionname,
      emailAddress: d.emailaddress,
      website: d.websiteurl,
      isDisabled: d.isdisabled,
      createdAt: d.createdon,
    };
  }

  async listBusinessUnits(params?: PaginationParams): Promise<PaginatedResponse<BusinessUnit>> {
    const limit = params?.limit || 20;
    const offset = params?.offset || 0;
    const queryParams = new URLSearchParams({
      $select: 'businessunitid,name,_parentbusinessunitid_value,divisionname,emailaddress,websiteurl,isdisabled,createdon',
      $top: String(limit),
      $skip: String(offset),
      $filter: 'isdisabled eq false',
      $orderby: 'name asc',
    });
    let url = `/businessunits?${queryParams}`;
    if (params?.cursor) url = params.cursor;
    const response = await this.request<DynamicsODataResponse<DynamicsBusinessUnit>>(url);
    return {
      items: response.value.map((b) => this.mapBusinessUnit(b)),
      count: response.value.length,
      total: response['@odata.count'],
      hasMore: !!response['@odata.nextLink'],
      nextCursor: response['@odata.nextLink'],
    };
  }

  async getBusinessUnit(id: string): Promise<BusinessUnit> {
    const bu = await this.request<DynamicsBusinessUnit>(`/businessunits(${id})`);
    return this.mapBusinessUnit(bu);
  }

  // ===========================================================================
  // Dynamics-specific: FetchXML Query
  // ===========================================================================

  async executeFetchXml(entity: string, fetchXml: string): Promise<unknown[]> {
    const encodedFetchXml = encodeURIComponent(fetchXml);
    const response = await this.request<DynamicsODataResponse<unknown>>(
      `/${entity}s?fetchXml=${encodedFetchXml}`
    );
    return response.value;
  }

  // ===========================================================================
  // Batch Operations
  // ===========================================================================

  async batchCreate(entityType: string, records: Record<string, unknown>[], continueOnError = false): Promise<{
    succeeded: number;
    failed: number;
    errors: { index: number; message: string }[];
    createdIds: string[];
  }> {
    const results = { succeeded: 0, failed: 0, errors: [] as { index: number; message: string }[], createdIds: [] as string[] };

    // Dynamics 365 batch requests use $batch endpoint
    // For simplicity, we'll process in parallel with Promise.allSettled
    const promises = records.map(async (record, index) => {
      try {
        const { id } = await this.createEntity(`/${entityType}s`, record);
        return { index, id, success: true };
      } catch (error) {
        return { index, error: error instanceof Error ? error.message : 'Unknown error', success: false };
      }
    });

    const settled = await Promise.allSettled(promises);
    for (const result of settled) {
      if (result.status === 'fulfilled') {
        if (result.value.success) {
          results.succeeded++;
          results.createdIds.push(result.value.id!);
        } else {
          results.failed++;
          results.errors.push({ index: result.value.index, message: result.value.error! });
          if (!continueOnError) break;
        }
      }
    }
    return results;
  }

  async batchUpdate(entityType: string, updates: { id: string; data: Record<string, unknown> }[], continueOnError = false): Promise<{
    succeeded: number;
    failed: number;
    errors: { index: number; message: string }[];
  }> {
    const results = { succeeded: 0, failed: 0, errors: [] as { index: number; message: string }[] };

    const promises = updates.map(async (update, index) => {
      try {
        await this.request(`/${entityType}s(${update.id})`, {
          method: 'PATCH',
          body: JSON.stringify(update.data),
        });
        return { index, success: true };
      } catch (error) {
        return { index, error: error instanceof Error ? error.message : 'Unknown error', success: false };
      }
    });

    const settled = await Promise.allSettled(promises);
    for (const result of settled) {
      if (result.status === 'fulfilled') {
        if (result.value.success) {
          results.succeeded++;
        } else {
          results.failed++;
          results.errors.push({ index: result.value.index, message: result.value.error! });
          if (!continueOnError) break;
        }
      }
    }
    return results;
  }

  async batchDelete(entityType: string, ids: string[], continueOnError = false): Promise<{
    succeeded: number;
    failed: number;
    errors: { index: number; message: string }[];
  }> {
    const results = { succeeded: 0, failed: 0, errors: [] as { index: number; message: string }[] };

    const promises = ids.map(async (id, index) => {
      try {
        await this.request(`/${entityType}s(${id})`, { method: 'DELETE' });
        return { index, success: true };
      } catch (error) {
        return { index, error: error instanceof Error ? error.message : 'Unknown error', success: false };
      }
    });

    const settled = await Promise.allSettled(promises);
    for (const result of settled) {
      if (result.status === 'fulfilled') {
        if (result.value.success) {
          results.succeeded++;
        } else {
          results.failed++;
          results.errors.push({ index: result.value.index, message: result.value.error! });
          if (!continueOnError) break;
        }
      }
    }
    return results;
  }

  async batchUpsert(entityType: string, records: { alternateKey?: Record<string, string>; id?: string; data: Record<string, unknown> }[], continueOnError = false): Promise<{
    succeeded: number;
    failed: number;
    created: number;
    updated: number;
    errors: { index: number; message: string }[];
  }> {
    const results = { succeeded: 0, failed: 0, created: 0, updated: 0, errors: [] as { index: number; message: string }[] };

    for (let index = 0; index < records.length; index++) {
      const record = records[index];
      try {
        let url = `/${entityType}s`;
        if (record.id) {
          url += `(${record.id})`;
        } else if (record.alternateKey) {
          const keyParts = Object.entries(record.alternateKey).map(([k, v]) => `${k}='${v}'`).join(',');
          url += `(${keyParts})`;
        }

        await this.request<unknown>(url, {
          method: record.id || record.alternateKey ? 'PATCH' : 'POST',
          body: JSON.stringify(record.data),
          headers: record.id || record.alternateKey ? { 'If-Match': '*' } : {},
        });

        results.succeeded++;
        if (record.id || record.alternateKey) {
          results.updated++;
        } else {
          results.created++;
        }
      } catch (error) {
        results.failed++;
        results.errors.push({ index, message: error instanceof Error ? error.message : 'Unknown error' });
        if (!continueOnError) break;
      }
    }
    return results;
  }

  // ===========================================================================
  // OData Query
  // ===========================================================================

  async executeQuery(params: {
    entityType: string;
    select?: string;
    filter?: string;
    orderby?: string;
    expand?: string;
    top?: number;
    skip?: number;
    count?: boolean;
  }): Promise<{ records: unknown[]; totalCount?: number; hasMore: boolean }> {
    const queryParams = new URLSearchParams();
    if (params.select) queryParams.set('$select', params.select);
    if (params.filter) queryParams.set('$filter', params.filter);
    if (params.orderby) queryParams.set('$orderby', params.orderby);
    if (params.expand) queryParams.set('$expand', params.expand);
    if (params.top) queryParams.set('$top', String(params.top));
    if (params.skip) queryParams.set('$skip', String(params.skip));
    if (params.count) queryParams.set('$count', 'true');

    const response = await this.request<DynamicsODataResponse<unknown>>(`/${params.entityType}s?${queryParams}`);
    return {
      records: response.value,
      totalCount: response['@odata.count'],
      hasMore: !!response['@odata.nextLink'],
    };
  }

  async executeAggregate(params: {
    entityType: string;
    aggregateFunction: 'count' | 'sum' | 'avg' | 'min' | 'max';
    field?: string;
    filter?: string;
    groupBy?: string;
  }): Promise<unknown> {
    // Build FetchXML for aggregate query
    let aggregate = '';
    if (params.aggregateFunction === 'count') {
      aggregate = `<attribute name="${params.field || params.entityType + 'id'}" aggregate="count" alias="count"/>`;
    } else {
      aggregate = `<attribute name="${params.field}" aggregate="${params.aggregateFunction}" alias="result"/>`;
    }

    let fetchXml = `<fetch aggregate="true"><entity name="${params.entityType}">${aggregate}`;
    if (params.filter) {
      fetchXml += `<filter><condition attribute="statecode" operator="eq" value="0"/></filter>`;
    }
    if (params.groupBy) {
      fetchXml += `<attribute name="${params.groupBy}" groupby="true" alias="groupby"/>`;
    }
    fetchXml += `</entity></fetch>`;

    return this.executeFetchXml(params.entityType, fetchXml);
  }

  async getRecordCount(entityType: string, filter?: string): Promise<number> {
    const queryParams = new URLSearchParams({
      $count: 'true',
      $top: '0',
    });
    if (filter) queryParams.set('$filter', filter);

    const response = await this.request<DynamicsODataResponse<unknown>>(`/${entityType}s?${queryParams}`);
    return response['@odata.count'] || 0;
  }

  async executeSavedQuery(savedQueryId: string, entityType: string, top = 50): Promise<unknown[]> {
    const query = await this.request<{ fetchxml: string }>(`/savedqueries(${savedQueryId})?$select=fetchxml`);
    return this.executeFetchXml(entityType, query.fetchxml);
  }

  async executeUserQuery(userQueryId: string, entityType: string, top = 50): Promise<unknown[]> {
    const query = await this.request<{ fetchxml: string }>(`/userqueries(${userQueryId})?$select=fetchxml`);
    return this.executeFetchXml(entityType, query.fetchxml);
  }

  // ===========================================================================
  // Metadata
  // ===========================================================================

  async listEntityMetadata(params?: { filter?: string; includeCustom?: boolean }): Promise<{
    logicalName: string;
    displayName: string;
    description?: string;
    isCustomEntity: boolean;
    primaryIdAttribute: string;
    primaryNameAttribute: string;
  }[]> {
    let url = '/EntityDefinitions?$select=LogicalName,DisplayName,Description,IsCustomEntity,PrimaryIdAttribute,PrimaryNameAttribute';
    const filters: string[] = [];
    if (params?.filter) {
      filters.push(`contains(LogicalName,'${params.filter}')`);
    }
    if (params?.includeCustom === false) {
      filters.push('IsCustomEntity eq false');
    }
    if (filters.length > 0) {
      url += `&$filter=${filters.join(' and ')}`;
    }

    const response = await this.request<{ value: Array<{
      LogicalName: string;
      DisplayName: { UserLocalizedLabel?: { Label: string } };
      Description: { UserLocalizedLabel?: { Label: string } };
      IsCustomEntity: boolean;
      PrimaryIdAttribute: string;
      PrimaryNameAttribute: string;
    }> }>(url);

    return response.value.map(e => ({
      logicalName: e.LogicalName,
      displayName: e.DisplayName?.UserLocalizedLabel?.Label || e.LogicalName,
      description: e.Description?.UserLocalizedLabel?.Label,
      isCustomEntity: e.IsCustomEntity,
      primaryIdAttribute: e.PrimaryIdAttribute,
      primaryNameAttribute: e.PrimaryNameAttribute,
    }));
  }

  async getEntityMetadata(entityLogicalName: string): Promise<unknown> {
    return this.request(`/EntityDefinitions(LogicalName='${entityLogicalName}')`);
  }

  async listEntityAttributes(entityLogicalName: string, attributeType?: string): Promise<{
    logicalName: string;
    displayName: string;
    attributeType: string;
    isRequired: boolean;
    isCustomAttribute: boolean;
    description?: string;
  }[]> {
    let url = `/EntityDefinitions(LogicalName='${entityLogicalName}')/Attributes?$select=LogicalName,DisplayName,AttributeType,RequiredLevel,IsCustomAttribute,Description`;
    if (attributeType) {
      url += `&$filter=AttributeType eq Microsoft.Dynamics.CRM.AttributeTypeCode'${attributeType}'`;
    }

    const response = await this.request<{ value: Array<{
      LogicalName: string;
      DisplayName: { UserLocalizedLabel?: { Label: string } };
      AttributeType: string;
      RequiredLevel: { Value: string };
      IsCustomAttribute: boolean;
      Description: { UserLocalizedLabel?: { Label: string } };
    }> }>(url);

    return response.value.map(a => ({
      logicalName: a.LogicalName,
      displayName: a.DisplayName?.UserLocalizedLabel?.Label || a.LogicalName,
      attributeType: a.AttributeType,
      isRequired: a.RequiredLevel?.Value === 'ApplicationRequired' || a.RequiredLevel?.Value === 'SystemRequired',
      isCustomAttribute: a.IsCustomAttribute,
      description: a.Description?.UserLocalizedLabel?.Label,
    }));
  }

  async getAttributeMetadata(entityLogicalName: string, attributeLogicalName: string): Promise<unknown> {
    return this.request(`/EntityDefinitions(LogicalName='${entityLogicalName}')/Attributes(LogicalName='${attributeLogicalName}')`);
  }

  async getOptionSetValues(entityLogicalName: string, attributeLogicalName: string): Promise<{ value: number; label: string }[]> {
    const response = await this.request<{
      OptionSet?: { Options: Array<{ Value: number; Label: { UserLocalizedLabel?: { Label: string } } }> };
    }>(`/EntityDefinitions(LogicalName='${entityLogicalName}')/Attributes(LogicalName='${attributeLogicalName}')/Microsoft.Dynamics.CRM.PicklistAttributeMetadata?$expand=OptionSet`);

    if (!response.OptionSet?.Options) return [];
    return response.OptionSet.Options.map(o => ({
      value: o.Value,
      label: o.Label?.UserLocalizedLabel?.Label || String(o.Value),
    }));
  }

  async getGlobalOptionSet(optionSetName: string): Promise<{ value: number; label: string }[]> {
    const response = await this.request<{
      Options: Array<{ Value: number; Label: { UserLocalizedLabel?: { Label: string } } }>;
    }>(`/GlobalOptionSetDefinitions(Name='${optionSetName}')`);

    if (!response.Options) return [];
    return response.Options.map(o => ({
      value: o.Value,
      label: o.Label?.UserLocalizedLabel?.Label || String(o.Value),
    }));
  }

  // ===========================================================================
  // Relationships
  // ===========================================================================

  async associateRecords(sourceEntityType: string, sourceId: string, targetEntityType: string, targetId: string, relationshipName: string): Promise<void> {
    await this.request(`/${sourceEntityType}s(${sourceId})/${relationshipName}/$ref`, {
      method: 'POST',
      body: JSON.stringify({ '@odata.id': `${this.baseUrl}/${targetEntityType}s(${targetId})` }),
    });
  }

  async disassociateRecords(sourceEntityType: string, sourceId: string, targetEntityType: string, targetId: string, relationshipName: string): Promise<void> {
    await this.request(`/${sourceEntityType}s(${sourceId})/${relationshipName}(${targetId})/$ref`, {
      method: 'DELETE',
    });
  }

  async listRelatedRecords(entityType: string, id: string, navigationProperty: string, params?: { select?: string; limit?: number }): Promise<unknown[]> {
    const queryParams = new URLSearchParams();
    if (params?.select) queryParams.set('$select', params.select);
    if (params?.limit) queryParams.set('$top', String(params.limit));

    const response = await this.request<DynamicsODataResponse<unknown>>(
      `/${entityType}s(${id})/${navigationProperty}?${queryParams}`
    );
    return response.value;
  }
}

// =============================================================================
// Factory Function
// =============================================================================

/**
 * Create a Dynamics 365 client instance with tenant-specific credentials.
 *
 * MULTI-TENANT: Each request provides its own credentials via headers,
 * allowing a single server deployment to serve multiple tenants.
 *
 * @param credentials - Tenant credentials parsed from request headers
 */
export function createCrmClient(credentials: TenantCredentials): CrmClient {
  return new DynamicsClientImpl(credentials);
}
