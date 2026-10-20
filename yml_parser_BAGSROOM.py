#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
import sys
import time
import requests
import xml.etree.ElementTree as ET
import hashlib
from datetime import datetime
import logging
import json
from typing import List, Dict, Any, Optional, Tuple, Set
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler("feed_processor.log", encoding='utf-8'),
        logging.StreamHandler(sys.stdout)
    ]
)
logger = logging.getLogger("feed_processor")

# Constants
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
CREDENTIALS_FILE = 'key_sheet.json'
PRODUCT_SHEET_NAME = 'Export Products Sheet'
CATEGORY_SHEET_NAME = 'Export Groups Sheet'
BATCH_SIZE = 200  # Number of rows to process in a batch
API_DELAY = 1.05  # Delay in seconds to respect API rate limit (approximately 60 requests/minute)


class FeedProcessor:
    def __init__(self, spreadsheet_id: str, feeds: List[str]):
        """
        Initialize the Feed Processor.

        Args:
            spreadsheet_id: The ID of the Google Sheet
            feeds: List of feed URLs to process
        """
        self.spreadsheet_id = spreadsheet_id
        self.feeds = feeds
        self.service = self._get_sheets_service()

        # Store categories and products from all feeds
        self.categories = {}  # id -> category data
        self.products = {}  # id -> product data

    def _get_sheets_service(self):
        """Create and return Google Sheets API service."""
        try:
            creds = Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=SCOPES)
            service = build('sheets', 'v4', credentials=creds)
            return service.spreadsheets()
        except Exception as e:
            logger.error(f"Failed to create Google Sheets service: {e}")
            raise

    def fetch_and_parse_feed(self, feed_url: str) -> Tuple[Dict[str, Dict], Dict[str, Dict]]:
        """
        Fetch and parse the XML feed.

        Args:
            feed_url: URL of the feed to fetch

        Returns:
            Tuple of (categories_dict, products_dict)
        """
        logger.info(f"Fetching feed from: {feed_url}")

        try:
            response = requests.get(feed_url, timeout=30)
            response.raise_for_status()

            # Parse XML content
            root = ET.fromstring(response.content)

            # Parse categories
            categories_dict = {}
            for category in root.findall('.//categories/category'):
                category_id = category.get('id')
                parent_id = category.get('parentId', '')
                name = category.text

                categories_dict[category_id] = {
                    'id': category_id,
                    'parentId': parent_id,
                    'name': name
                }

            # Parse products
            products_dict = {}
            for offer in root.findall('.//offers/offer'):
                product_id = offer.get('id')
                available = offer.get('available', 'false')
                group_id = offer.get('group_id', '')

                # Initialize product data with critical fields
                product_data = {
                    'id': product_id,
                    'available': available,
                    'group_id': group_id,
                    'pictures': []
                }

                # Process direct child elements
                for child in offer:
                    tag = child.tag

                    # Handle pictures separately
                    if tag == 'picture':
                        product_data['pictures'].append(child.text)
                    else:
                        product_data[tag] = child.text

                # Process parameters
                for param in offer.findall('./param'):
                    param_name = param.get('name', '')
                    if param_name:
                        product_data[f"param_{param_name}"] = param.text

                # Ensure critical fields exist and are properly formatted
                # Convert 'available' to lowercase boolean string
                product_data['available'] = str(product_data.get('available', 'false')).lower()

                # Ensure 'price' is a number if it exists
                if 'price' in product_data:
                    try:
                        product_data['price'] = str(float(product_data['price']))
                    except (ValueError, TypeError):
                        logger.warning(f"Invalid price format for product {product_id}: {product_data.get('price')}")
                        product_data['price'] = '0'

                # Ensure 'oldprice' is a number if it exists
                if 'oldprice' in product_data:
                    try:
                        product_data['oldprice'] = str(float(product_data['oldprice']))
                    except (ValueError, TypeError):
                        logger.warning(
                            f"Invalid oldprice format for product {product_id}: {product_data.get('oldprice')}")
                        product_data['oldprice'] = '0'

                # Ensure 'quantity_in_stock' is an integer if it exists
                if 'quantity_in_stock' in product_data:
                    try:
                        product_data['quantity_in_stock'] = str(int(float(product_data['quantity_in_stock'])))
                    except (ValueError, TypeError):
                        logger.warning(
                            f"Invalid quantity format for product {product_id}: {product_data.get('quantity_in_stock')}")
                        product_data['quantity_in_stock'] = '0'

                products_dict[product_id] = product_data

            logger.info(f"Parsed {len(categories_dict)} categories and {len(products_dict)} products from feed")
            return categories_dict, products_dict

        except requests.exceptions.RequestException as e:
            logger.error(f"Error fetching feed: {e}")
            return {}, {}
        except ET.ParseError as e:
            logger.error(f"Error parsing XML feed: {e}")
            return {}, {}
        except Exception as e:
            logger.error(f"Unexpected error processing feed: {e}")
            return {}, {}

    def process_all_feeds(self):
        """Process all feeds and combine the data, handling duplicates across feeds."""
        all_products = {}
        all_categories = {}

        # First pass: collect all data from all feeds
        for feed_url in self.feeds:
            categories, products = self.fetch_and_parse_feed(feed_url)

            # Detect duplicate products across feeds by creating a hash of critical fields
            for product_id, product_data in products.items():
                # Check if product already exists in our collection
                if product_id in all_products:
                    # Create hash of critical fields to check if they're functionally the same
                    current_hash = self._get_product_hash(all_products[product_id])
                    new_hash = self._get_product_hash(product_data)

                    if current_hash != new_hash:
                        logger.warning(f"Duplicate product ID {product_id} found with different data across feeds!")

                        # Prioritize the product with more data or non-empty critical fields
                        if self._product_has_better_data(product_data, all_products[product_id]):
                            logger.info(f"Using product data from {feed_url} for ID {product_id} as it has better data")
                            all_products[product_id] = product_data

                    # If hashes match, they're functionally the same - keep what we have
                else:
                    # New product, add it
                    all_products[product_id] = product_data

            # Categories are simpler, just merge them
            all_categories.update(categories)

        # Set the merged data to our instance variables
        self.categories = all_categories
        self.products = all_products

        logger.info(
            f"Processed a total of {len(self.categories)} unique categories and {len(self.products)} unique products")

    def _get_product_hash(self, product_data: Dict) -> str:
        """Generate a hash of critical product fields to compare products."""
        # Select critical fields for comparison
        critical_fields = ['available', 'price', 'oldprice', 'quantity_in_stock', 'name', 'categoryId']

        # Build a string of these fields
        hash_string = ""
        for field in critical_fields:
            if field in product_data:
                hash_string += f"{field}:{product_data[field]}|"

        # Create and return hash
        return hashlib.md5(hash_string.encode()).hexdigest()

    def _product_has_better_data(self, product1: Dict, product2: Dict) -> bool:
        """
        Determine if product1 has better data than product2.
        Better data means:
        1. Has critical fields that product2 doesn't have
        2. Has non-empty critical fields where product2 has empty ones
        3. Has more fields overall
        """
        critical_fields = ['available', 'price', 'oldprice', 'quantity_in_stock']

        # Check critical fields first
        for field in critical_fields:
            # product1 has field that product2 doesn't
            if field in product1 and field not in product2:
                return True

            # Both have field but product2's is empty while product1's isn't
            if (field in product1 and field in product2 and
                    product1[field] and not product2[field]):
                return True

        # If tied on critical fields, check total field count
        return len(product1) > len(product2)

    def get_sheet_data(self, sheet_name: str) -> List[List]:
        """
        Get all data from a sheet.

        Args:
            sheet_name: Name of the sheet to read

        Returns:
            List of rows (each row is a list of cell values)
        """
        try:
            result = self.service.values().get(
                spreadsheetId=self.spreadsheet_id,
                range=sheet_name
            ).execute()

            values = result.get('values', [])
            logger.info(f"Retrieved {len(values)} rows from {sheet_name}")
            return values
        except HttpError as e:
            if e.resp.status == 400:
                logger.info(f"Sheet {sheet_name} doesn't exist yet or is empty")
                return []
            else:
                logger.error(f"Error retrieving data from sheet {sheet_name}: {e}")
                raise

    def create_or_update_sheet_headers(self, sheet_name: str, headers: List[str]):
        """
        Ensure the sheet exists and has the correct headers.

        Args:
            sheet_name: Name of the sheet
            headers: List of header names
        """
        try:
            # Check if sheet exists
            sheet_metadata = self.service.get(spreadsheetId=self.spreadsheet_id).execute()
            sheet_exists = False

            for sheet in sheet_metadata.get('sheets', []):
                if sheet['properties']['title'] == sheet_name:
                    sheet_exists = True
                    break

            if not sheet_exists:
                # Create sheet if it doesn't exist
                request = {
                    'addSheet': {
                        'properties': {
                            'title': sheet_name
                        }
                    }
                }
                self.service.batchUpdate(
                    spreadsheetId=self.spreadsheet_id,
                    body={'requests': [request]}
                ).execute()
                logger.info(f"Created new sheet: {sheet_name}")

                # Add headers to new sheet
                self.service.values().update(
                    spreadsheetId=self.spreadsheet_id,
                    range=f"{sheet_name}!A1",
                    valueInputOption="USER_ENTERED",
                    body={
                        'values': [headers]
                    }
                ).execute()
                logger.info(f"Added headers to {sheet_name}")
            else:
                # Check existing headers and update if needed
                existing_data = self.get_sheet_data(sheet_name)

                if not existing_data:
                    # Sheet exists but is empty
                    self.service.values().update(
                        spreadsheetId=self.spreadsheet_id,
                        range=f"{sheet_name}!A1",
                        valueInputOption="USER_ENTERED",
                        body={
                            'values': [headers]
                        }
                    ).execute()
                    logger.info(f"Added headers to empty sheet {sheet_name}")
                else:
                    existing_headers = existing_data[0]

                    # Check if we need to add new headers
                    new_headers = [h for h in headers if h not in existing_headers]
                    if new_headers:
                        # Append new headers to existing ones
                        updated_headers = existing_headers + new_headers
                        self.service.values().update(
                            spreadsheetId=self.spreadsheet_id,
                            range=f"{sheet_name}!A1",
                            valueInputOption="USER_ENTERED",
                            body={
                                'values': [updated_headers]
                            }
                        ).execute()
                        logger.info(f"Updated headers in {sheet_name}, added: {new_headers}")
                        return updated_headers
                    return existing_headers

            self._format_headers_bold(sheet_name)
            return headers

        except HttpError as e:
            logger.error(f"Error updating sheet headers: {e}")
            raise

    def prepare_categories_data(self, existing_headers: List[str]) -> List[List]:
        """
        Prepare category data for batch update.

        Args:
            existing_headers: List of column headers in the sheet

        Returns:
            List of rows with category data
        """
        rows = []

        for category_id, category_data in self.categories.items():
            row = [""] * len(existing_headers)

            for idx, header in enumerate(existing_headers):
                if header in category_data:
                    row[idx] = category_data[header]

            rows.append(row)

        return rows

    def prepare_products_data(self, existing_headers: List[str]) -> List[List]:
        """
        Prepare product data for batch update.

        Args:
            existing_headers: List of column headers in the sheet

        Returns:
            List of rows with product data
        """
        rows = []

        # Define critical fields that must be included
        critical_fields = ['available', 'price', 'oldprice', 'quantity_in_stock']

        for product_id, product_data in self.products.items():
            row = [""] * len(existing_headers)

            # First pass: fill in all existing data
            for idx, header in enumerate(existing_headers):
                if header == 'pictures':
                    # Join multiple picture URLs with delimiter
                    if 'pictures' in product_data and product_data['pictures']:
                        row[idx] = ", ".join(product_data['pictures'])
                elif header in product_data:
                    row[idx] = product_data[header]

            # Second pass: ensure critical fields are properly handled
            for field in critical_fields:
                try:
                    idx = existing_headers.index(field)
                    # If critical field is missing in data, log warning
                    if not row[idx] and field in product_data:
                        logger.warning(f"Critical field {field} might be missing or empty for product {product_id}")

                    # For available field, ensure it's a proper boolean string
                    if field == 'available' and field in product_data:
                        row[idx] = str(product_data[field]).lower()

                    # For numeric fields, ensure they're proper numbers
                    if field in ['price', 'oldprice', 'quantity_in_stock'] and field in product_data:
                        try:
                            # Try to convert to float/int to validate
                            if field == 'quantity_in_stock':
                                row[idx] = str(int(float(product_data[field])))
                            else:
                                row[idx] = str(float(product_data[field]))
                        except (ValueError, TypeError):
                            logger.warning(
                                f"Invalid numeric value for {field} in product {product_id}: {product_data[field]}")
                except ValueError:
                    # Header doesn't exist in sheet
                    pass

            rows.append(row)

        return rows

    def delete_items_not_in_feed(self, sheet_name: str, id_column: str, current_ids: Set[str]):
        """
        Delete rows for items that are no longer in the feed.

        Args:
            sheet_name: Name of the sheet
            id_column: Name of the ID column
            current_ids: Set of IDs from the current feed
        """
        try:
            sheet_data = self.get_sheet_data(sheet_name)
            if not sheet_data:
                return

            headers = sheet_data[0]
            try:
                id_index = headers.index(id_column)
            except ValueError:
                logger.error(f"ID column '{id_column}' not found in sheet {sheet_name}")
                return

            rows_to_delete = []

            # Find rows to delete (skip header row)
            for i, row in enumerate(sheet_data[1:], start=1):
                if len(row) > id_index:
                    row_id = row[id_index]
                    if row_id and row_id not in current_ids:
                        rows_to_delete.append(i + 1)  # +1 for 1-indexed rows in API

            # Delete rows in reverse order to avoid shifting indices
            if rows_to_delete:
                rows_to_delete.sort(reverse=True)

                # Delete in batches to respect API limits
                batch_size = 20  # Smaller batch for delete operations
                for i in range(0, len(rows_to_delete), batch_size):
                    batch = rows_to_delete[i:i + batch_size]
                    requests = []

                    for row_index in batch:
                        requests.append({
                            'deleteDimension': {
                                'range': {
                                    'sheetId': self._get_sheet_id(sheet_name),
                                    'dimension': 'ROWS',
                                    'startIndex': row_index - 1,  # Convert to 0-indexed
                                    'endIndex': row_index  # End is exclusive
                                }
                            }
                        })

                    if requests:
                        self.service.batchUpdate(
                            spreadsheetId=self.spreadsheet_id,
                            body={'requests': requests}
                        ).execute()

                        # Respect API rate limit
                        time.sleep(API_DELAY)

                logger.info(f"Deleted {len(rows_to_delete)} rows from {sheet_name} for items no longer in feed")

        except HttpError as e:
            logger.error(f"Error deleting rows: {e}")
            raise

    def _get_sheet_id(self, sheet_name: str) -> int:
        """Get the sheet ID for a given sheet name."""
        sheet_metadata = self.service.get(spreadsheetId=self.spreadsheet_id).execute()
        for sheet in sheet_metadata.get('sheets', []):
            if sheet['properties']['title'] == sheet_name:
                return sheet['properties']['sheetId']
        raise ValueError(f"Sheet {sheet_name} not found")

    def update_sheet_in_batches(self, sheet_name: str, data: List[List], id_column: str):
        """
        Update a sheet with data in batches, handling duplicates.

        Args:
            sheet_name: Name of the sheet
            data: List of rows to update
            id_column: Name of the ID column to identify existing records
        """
        if not data:
            logger.info(f"No data to update for {sheet_name}")
            return

        try:
            # Get existing data to determine what needs to be updated vs. appended
            existing_data = self.get_sheet_data(sheet_name)

            if not existing_data:
                # Sheet is empty or doesn't have data, just append everything
                self._batch_append_rows(sheet_name, data)
                return

            headers = existing_data[0]
            try:
                id_index = headers.index(id_column)
            except ValueError:
                logger.error(f"ID column '{id_column}' not found in {sheet_name}")
                return

            # Create a map of existing IDs to row indices
            # This will detect duplicates in the sheet
            existing_ids = {}
            duplicate_ids = set()

            for i, row in enumerate(existing_data[1:], start=1):  # Skip header row
                if len(row) > id_index and row[id_index]:
                    row_id = row[id_index]
                    if row_id in existing_ids:
                        # This is a duplicate
                        duplicate_ids.add(row_id)
                        logger.info(f"Found duplicate ID {row_id} in {sheet_name}")
                    existing_ids[row_id] = i

            # Handle duplicates first - remove all but one instance
            if duplicate_ids:
                logger.info(f"Found {len(duplicate_ids)} duplicate IDs in {sheet_name}, cleaning up...")
                rows_to_delete = []

                # For each duplicate ID, keep only the first occurrence
                for dup_id in duplicate_ids:
                    # Find all rows with this ID
                    dup_rows = [i for i, row in enumerate(existing_data[1:], start=1)
                                if len(row) > id_index and row[id_index] == dup_id]

                    # Keep the first one, delete the rest
                    rows_to_delete.extend(dup_rows[1:])

                # Delete duplicate rows
                if rows_to_delete:
                    rows_to_delete.sort(reverse=True)  # Delete from bottom up

                    # Delete in batches
                    batch_size = 20
                    for i in range(0, len(rows_to_delete), batch_size):
                        batch = rows_to_delete[i:i + batch_size]
                        requests = []

                        for row_index in batch:
                            requests.append({
                                'deleteDimension': {
                                    'range': {
                                        'sheetId': self._get_sheet_id(sheet_name),
                                        'dimension': 'ROWS',
                                        'startIndex': row_index,  # Row index is already 1-based
                                        'endIndex': row_index + 1  # End is exclusive
                                    }
                                }
                            })

                        if requests:
                            self.service.batchUpdate(
                                spreadsheetId=self.spreadsheet_id,
                                body={'requests': requests}
                            ).execute()

                            # Respect API rate limit
                            time.sleep(API_DELAY)

                    # Re-fetch data after deletion
                    existing_data = self.get_sheet_data(sheet_name)
                    headers = existing_data[0]

                    # Rebuild existing_ids map
                    existing_ids = {}
                    for i, row in enumerate(existing_data[1:], start=1):
                        if len(row) > id_index and row[id_index]:
                            existing_ids[row[id_index]] = i

                    logger.info(f"Removed {len(rows_to_delete)} duplicate rows from {sheet_name}")

            # Now handle regular updates and inserts
            updates = []
            inserts = []

            for row in data:
                if len(row) > id_index and row[id_index] in existing_ids:
                    # This is an update
                    updates.append((existing_ids[row[id_index]], row))
                else:
                    # This is a new insert
                    inserts.append(row)

            # Process updates in batches
            if updates:
                self._batch_update_rows(sheet_name, updates, headers)

            # Process inserts in batches
            if inserts:
                self._batch_append_rows(sheet_name, inserts)

        except HttpError as e:
            logger.error(f"Error updating sheet in batches: {e}")
            raise

    def _batch_update_rows(self, sheet_name: str, updates: List[Tuple[int, List]], headers: List[str]):
        """
        Update existing rows in batches, with special handling for critical fields.

        Args:
            sheet_name: Name of the sheet
            updates: List of (row_index, row_data) tuples
            headers: Column headers
        """
        # Only check critical fields for products sheet
        if sheet_name == PRODUCT_SHEET_NAME:
            # Define critical fields that must be updated
            critical_fields = ['available', 'price', 'oldprice', 'quantity_in_stock']

            # Find indices of critical fields in headers
            critical_indices = {}
            for field in critical_fields:
                try:
                    critical_indices[field] = headers.index(field)
                except ValueError:
                    # This critical field isn't in headers
                    pass

            # Get the existing data for critical field comparison
            existing_data = self.get_sheet_data(sheet_name)
        else:
            critical_indices = None
            existing_data = None

        # Group updates into batches
        batches = [updates[i:i + BATCH_SIZE] for i in range(0, len(updates), BATCH_SIZE)]

        # Process each batch
        for batch_idx, batch in enumerate(batches):
            logger.info(f"Processing update batch {batch_idx + 1}/{len(batches)} for {sheet_name}")

            # Clean data before updating
            cleaned_batch = []
            for row_idx, row_data in batch:
                cleaned_row = []
                for value in row_data:
                    if isinstance(value, str):
                        # Remove extra whitespace and empty lines
                        value = value.strip()
                        value = '\n'.join(filter(None, [line.strip() for line in value.split('\n')]))
                    cleaned_row.append(value)
                cleaned_batch.append((row_idx, cleaned_row))
            batch = cleaned_batch

            # Create data for valuesBatchUpdate
            data = []
            for row_idx, row_data in batch:
                # For products sheet, check critical fields
                if sheet_name == PRODUCT_SHEET_NAME and critical_indices and row_idx < len(existing_data):
                    existing_row = existing_data[row_idx]

                    # Check each critical field
                    for field, idx in critical_indices.items():
                        if idx < len(row_data):
                            # If row_data has empty value for critical field but existing has value,
                            # log a warning
                            if (not row_data[idx] or row_data[idx] == '') and idx < len(existing_row) and existing_row[
                                idx]:
                                logger.warning(f"Critical field {field} has empty value in feed but exists in sheet. "
                                               f"Row: {row_idx + 1}, Current: {existing_row[idx]}, Feed: empty")

                            # If feed value differs from sheet, log the change
                            #elif idx < len(existing_row) and existing_row[idx] != row_data[idx]:
                                #logger.info(f"Updating critical field {field} in row {row_idx + 1}: "
                                            #f"{existing_row[idx]} -> {row_data[idx]}")

                # Create range for this row
                row_range = f"{sheet_name}!A{row_idx + 1}:{self._column_letter(len(headers))}{row_idx + 1}"

                # Create update for this row
                data.append({
                    'range': row_range,
                    'values': [row_data]
                })

            # Execute batch update
            if data:
                body = {
                    'valueInputOption': 'USER_ENTERED',
                    'data': data
                }

                self.service.values().batchUpdate(
                    spreadsheetId=self.spreadsheet_id,
                    body=body
                ).execute()

                # Respect API rate limit
                time.sleep(API_DELAY)

        logger.info(f"Updated {len(updates)} existing rows in {sheet_name}")

    def _batch_append_rows(self, sheet_name: str, rows: List[List]):
        """
        Append new rows in batches.

        Args:
            sheet_name: Name of the sheet
            rows: List of rows to append
        """
        # Group rows into batches
        batches = [rows[i:i + BATCH_SIZE] for i in range(0, len(rows), BATCH_SIZE)]

        # Process each batch
        for batch_idx, batch in enumerate(batches):
            logger.info(f"Processing append batch {batch_idx + 1}/{len(batches)} for {sheet_name}")

            # Clean data before appending
            cleaned_batch = []
            for row in batch:
                cleaned_row = []
                for value in row:
                    if isinstance(value, str):
                        # Remove extra whitespace and empty lines
                        value = value.strip()
                        value = '\n'.join(filter(None, [line.strip() for line in value.split('\n')]))
                    cleaned_row.append(value)
                cleaned_batch.append(cleaned_row)

            self.service.values().append(
                spreadsheetId=self.spreadsheet_id,
                range=f"{sheet_name}!A1",
                valueInputOption="USER_ENTERED",
                insertDataOption="INSERT_ROWS",
                body={
                    'values': cleaned_batch
                }
            ).execute()

            # Respect API rate limit
            time.sleep(API_DELAY)

        logger.info(f"Appended {len(rows)} new rows to {sheet_name}")

    def _column_letter(self, column_number: int) -> str:
        """Convert column number to letter (1=A, 2=B, etc.)."""
        result = ""
        while column_number > 0:
            column_number, remainder = divmod(column_number - 1, 26)
            result = chr(65 + remainder) + result
        return result

    def _format_headers_bold(self, sheet_name: str):
        """Make the first row (headers) bold."""
        try:
            requests = [{
                "repeatCell": {
                    "range": {
                        "sheetId": self._get_sheet_id(sheet_name),
                        "startRowIndex": 0,
                        "endRowIndex": 1
                    },
                    "fields": "userEnteredFormat.textFormat.bold",
                    "cell": {
                        "userEnteredFormat": {
                            "textFormat": {
                                "bold": True
                            }
                        }
                    }
                }
            }]
            self.service.batchUpdate(
                spreadsheetId=self.spreadsheet_id,
                body={'requests': requests}
            ).execute()
        except Exception as e:
            logger.warning(f"Couldn't format headers: {e}")

    def _clear_data_formatting(self, sheet_name: str):
        """Remove bold from all data rows (except header)."""
        try:
            requests = [{
                "repeatCell": {
                    "range": {
                        "sheetId": self._get_sheet_id(sheet_name),
                        "startRowIndex": 1  # Skip header
                    },
                    "fields": "userEnteredFormat.textFormat.bold",
                    "cell": {
                        "userEnteredFormat": {
                            "textFormat": {
                                "bold": False
                            }
                        }
                    }
                }
            }]
            self.service.batchUpdate(
                spreadsheetId=self.spreadsheet_id,
                body={'requests': requests}
            ).execute()
        except Exception as e:
            logger.warning(f"Couldn't clear formatting: {e}")

    def get_category_headers(self) -> List[str]:
        """Get headers for the category sheet based on available data."""
        all_keys = set()
        for category in self.categories.values():
            all_keys.update(category.keys())

        # Ensure id is first
        headers = ['id']
        for key in sorted(all_keys):
            if key != 'id':
                headers.append(key)

        return headers

    def get_product_headers(self) -> List[str]:
        """Get headers for the product sheet based on available data."""
        all_keys = set()
        for product in self.products.values():
            all_keys.update(product.keys())

        # Special handling for pictures
        if 'pictures' not in all_keys and any('pictures' in p for p in self.products.values()):
            all_keys.add('pictures')

        # Ensure id is first
        headers = ['id']
        for key in sorted(all_keys):
            if key != 'id':
                headers.append(key)

        return headers

    def update_sheets(self):
        """Update categories and products in the Google Sheets."""
        # Process categories
        category_headers = self.get_category_headers()
        existing_category_headers = self.create_or_update_sheet_headers(CATEGORY_SHEET_NAME, category_headers)
        category_data = self.prepare_categories_data(existing_category_headers)
        self.update_sheet_in_batches(CATEGORY_SHEET_NAME, category_data, 'id')
        self.delete_items_not_in_feed(CATEGORY_SHEET_NAME, 'id', set(self.categories.keys()))
        self._clear_data_formatting(CATEGORY_SHEET_NAME)

        # Process products
        product_headers = self.get_product_headers()
        existing_product_headers = self.create_or_update_sheet_headers(PRODUCT_SHEET_NAME, product_headers)
        product_data = self.prepare_products_data(existing_product_headers)
        self.update_sheet_in_batches(PRODUCT_SHEET_NAME, product_data, 'id')
        self.delete_items_not_in_feed(PRODUCT_SHEET_NAME, 'id', set(self.products.keys()))
        self._clear_data_formatting(PRODUCT_SHEET_NAME)

    def run(self):
        """Run the complete feed processing and sheet update."""
        start_time = datetime.now()
        logger.info(f"=== Feed processing started at {start_time} ===")

        try:
            # Process all feeds
            self.process_all_feeds()

            # Update Google Sheets
            self.update_sheets()

            end_time = datetime.now()
            duration = (end_time - start_time).total_seconds()
            logger.info(f"=== Feed processing completed at {end_time} ===")
            logger.info(f"Total duration: {duration:.2f} seconds")

        except Exception as e:
            logger.error(f"Feed processing failed: {e}")
            raise

if __name__ == "__main__":
    # Configuration
    # 1. Прячем ID таблицы
    SPREADSHEET_ID = os.environ.get("PARS_BAGSROOM_SPREADSHEET_ID", "ВАШ_ID_ТАБЛИЦЫ_ПОСТАВЩИКА")

    # 2. Прячем ссылку на фид
    bagsroom_url = os.environ.get("BAGSROOM_FEED_URL")

    if bagsroom_url:
        # Если секрет найден, подставляем его в список
        FEEDS = [bagsroom_url]
        logger.info("URL поставщика BAGSROOM успешно получен из секретов.")
    else:
        # Если секрета нет (запуск на компе), используем эту ссылку
        FEEDS = [
            "https://bagsroom.com.ua/products_feed.xml?"
        ]
        logger.warning("Предупреждение: Ссылка BAGSROOM не найдена в секретах, используем локальное значение.")

    processor = FeedProcessor(SPREADSHEET_ID, FEEDS)
    processor.run()
