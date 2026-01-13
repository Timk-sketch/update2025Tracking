// =====================================================
// 00_Config.gs — Global configuration + constants
// =====================================================

const PROPS = PropertiesService.getScriptProperties();

const BANNED_ARCHIVE_ID = PROPS.getProperty('BANNED_ARCHIVE_ID'); // optional; if not set, banned list ignored
const BANNED_SHEET_NAME_PRIMARY = 'BANNEDEmailList';
const BANNED_SHEET_NAME_FALLBACK = 'BannedEmailList';

// ====== RAW HEADERS ======
const SQUARESPACE_ORDER_HEADERS = [
  "Order ID", "Order Number", "Created On", "Modified On", "Channel", "Test Mode", "Customer Email",
  "Billing First Name", "Billing Last Name", "Billing Address1", "Billing Address2", "Billing City", "Billing State",
  "Billing Country Code", "Billing Postal Code", "Billing Phone", "Shipping First Name", "Shipping Last Name",
  "Shipping Address1", "Shipping Address2", "Shipping City", "Shipping State", "Shipping Country Code",
  "Shipping Postal Code", "Shipping Phone", "Fulfillment Status", "Internal Notes", "Subtotal Currency",
  "Subtotal Value", "Shipping Total Currency", "Shipping Total Value", "Discount Total Currency",
  "Discount Total Value", "Tax Total Currency", "Tax Total Value", "Refunded Total Currency", "Refunded Total Value",
  "Grand Total Currency", "Grand Total Value", "LineItem ID", "LineItem SKU", "LineItem Weight", "LineItem Width",
  "LineItem Length", "LineItem Height", "LineItem Product ID", "LineItem Product Name", "LineItem Quantity",
  "LineItem Unit Price Currency", "LineItem Unit Price Value", "LineItem Customizations", "LineItem Type",
  "Shipping Lines (All)", "Discount Lines (All)"
];

// Clean, non-duplicated Shopify headers (aligned to row writes)
const SHOPIFY_ORDER_HEADERS = [
  "Order ID",
  "Order Number",
  "Created At",
  "Processed At",
  "Updated At",
  "Financial Status",
  "Fulfillment Status",
  "Currency",
  "Total Price",
  "Subtotal Price",
  "Total Tax",
  "Total Discounts",
  "Current Total Price",
  "Current Total Discounts",
  "Total Refunds",
  "Test Order",
  "Customer Email",
  "Customer First Name",
  "Customer Last Name",
  "Billing Name",
  "Billing Address1",
  "Billing Address2",
  "Billing City",
  "Billing Province",
  "Billing Country",
  "Billing Zip",
  "Billing Phone",
  "Shipping Name",
  "Shipping Address1",
  "Shipping Address2",
  "Shipping City",
  "Shipping Province",
  "Shipping Country",
  "Shipping Zip",
  "Shipping Phone",
  "Lineitem ID",
  "Lineitem Name",
  "Lineitem Quantity",
  "Lineitem Price",
  "Lineitem SKU",
  "Lineitem Requires Shipping",
  "Lineitem Taxable",
  "Lineitem Fulfillment Status",
  "Tags",
  "Note",
  "Gateway",
  "Total Weight",
  "Discount Codes",
  "Shipping Method",
  "Created At (Local)",
  "Processed At (Local)"
];

// ====== SHEET NAMES ======
const CLEAN_OUTPUT_SHEET = 'All_Orders_Clean';

// IMPORTANT: These MUST match what Summary expects
const CLEAN_HEADERS = [
  "platform",
  "order_id",
  "order_number",
  "order_date",
  "customer_email_raw",
  "customer_email_norm",
  "customer_name",
  "product_name",
  "sku",
  "quantity",
  "unit_price",
  "line_revenue",
  "order_discount_total",
  "order_refund_total",
  "order_net_revenue",
  "currency",
  "financial_status",
  "fulfillment_status",
  "tags",
  "source_sheet"
];

const SUMMARY_SHEET = "Orders_Summary_Report";
const OUTREACH_CONTROLS_SHEET = "Customer_Outreach_Controls";
const OUTREACH_OUTPUT_SHEET = "Customer_Outreach_List";
const MARKETING_CONTROLS_SHEET = "Marketing_Controls";
