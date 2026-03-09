using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media.Imaging;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using SecurityBadgePrinter.Models;
using SecurityBadgePrinter.Services;
using SkiaSharp;
using System.Text;
using System.Threading;
using System.Windows.Input;

namespace SecurityBadgePrinter
{
    public partial class MainWindow : Window
    {
        private readonly GraphServiceClient _graph;
        private readonly AppConfig _config;
        private readonly PrinterService _printer;
        private readonly BadgeRenderer _renderer = new BadgeRenderer();

        private SKBitmap? _lastBadge;
        private User? _selectedUser;
        private const string AllowedDomain = "@siderise.com";
        private CancellationTokenSource? _filterCts;
        private CancellationTokenSource? _filterOptionsCts;
        private bool UseGroupWhitelist => _config.AzureAd.AllowedGroupIds != null && _config.AzureAd.AllowedGroupIds.Count > 0;

        public MainWindow()
        {
            InitializeComponent();
            _graph = App.GraphClient;
            _config = App.Config;
            _printer = new PrinterService(_config.Printer.Name);
        }

        private void OnPreviewImageLoaded(object sender, RoutedEventArgs e)
        {
            // Ensure single preview clamps after the image has measured
            ConstrainSinglePreviewToViewport();
        }

        private void ConstrainSinglePreviewToViewport()
        {
            try
            {
                if (PreviewScroller == null || SinglePreviewBox == null || SinglePreviewBox.Visibility != Visibility.Visible) return;
                var vw = PreviewScroller.ViewportWidth;
                var vh = PreviewScroller.ViewportHeight;
                // Account for the PreviewFrame border/padding and a tiny safety margin
                double framePadH = 0, framePadV = 0;
                if (PreviewFrame != null)
                {
                    // Include an extra safety margin to avoid any clipping at edges
                    framePadH = PreviewFrame.BorderThickness.Left + PreviewFrame.BorderThickness.Right + PreviewFrame.Padding.Left + PreviewFrame.Padding.Right + 12;
                    framePadV = PreviewFrame.BorderThickness.Top + PreviewFrame.BorderThickness.Bottom + PreviewFrame.Padding.Top + PreviewFrame.Padding.Bottom + 12;
                }
                if (vw > 0) SinglePreviewBox.MaxWidth = Math.Max(0, vw - framePadH);
                if (vh > 0) SinglePreviewBox.MaxHeight = Math.Max(0, vh - framePadV);
                // Disable horizontal scrollbar, allow vertical if needed to avoid clipping
                PreviewScroller.HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled;
                PreviewScroller.VerticalScrollBarVisibility = ScrollBarVisibility.Auto;
            }
            catch { }
        }

        private void OnPreviewAreaSizeChanged(object sender, SizeChangedEventArgs e)
        {
            try
            {
                if (PreviewScroller == null) return;
                ConstrainSinglePreviewToViewport();

                // If name list is visible (>6), fit the names scroller to viewport to drive WrapPanel columns
                if (PreviewNamesContainer != null && PreviewNamesContainer.Visibility == Visibility.Visible && PreviewNamesScroll != null)
                {
                    var vw = PreviewScroller.ViewportWidth;
                    var vh = PreviewScroller.ViewportHeight;
                    if (vh > 0) PreviewNamesScroll.Height = vh - 56; // leave space for header
                    PreviewScroller.HorizontalScrollBarVisibility = ScrollBarVisibility.Auto;
                    PreviewScroller.VerticalScrollBarVisibility = ScrollBarVisibility.Disabled;
                }

                // Constrain multi preview grid to viewport height with a safety margin to avoid bottom clipping
                if (PreviewItems != null && PreviewItems.Visibility == Visibility.Visible)
                {
                    var vh = PreviewScroller.ViewportHeight;
                    if (vh > 0)
                    {
                        PreviewItems.MaxHeight = Math.Max(0, vh - 20);
                    }
                }
            }
            catch { }

        }

        private async void OnWindowLoaded(object sender, RoutedEventArgs e)
        {
            // Populate filters and load default user list (domain-restricted)
            try
            {
                StatusText.Text = "Loading directory...";
                await LoadFilterValuesAsync();
                await QueryAndBindUsersAsync(string.Empty);
                StatusText.Text = "Ready";

                // Hook editable text changes only for Office (to avoid fighting user selections)
                AttachEditableTextWatcher(OfficeBox as ComboBox);

                // Populate right banner info (app version and signed-in user)
                try
                {
                    var ver = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;
                    if (VersionText != null && ver != null)
                    {
                        VersionText.Text = $"v{ver.ToString(3)}"; // show Major.Minor.Build only
                    }
                }
                catch { }

                // No user in banner by request; only version is shown on the right
            }
            catch (Exception ex)
            {
                StatusText.Text = $"Load error: {ex.Message}";
            }
        }

        private void AttachEditableTextWatcher(ComboBox? cb)
        {
            if (cb == null || !cb.IsEditable) return;
            cb.ApplyTemplate();
            if (cb.Template.FindName("PART_EditableTextBox", cb) is TextBox tb)
            {
                tb.TextChanged += OnFilterTextChanged;
            }
        }

        private async Task LoadFilterValuesAsync()
        {
            List<User> users;
            List<string> officeOptions = new List<string>();
            
            if (UseGroupWhitelist)
            {
                var union = new Dictionary<string, User>(StringComparer.OrdinalIgnoreCase);
                
                // Fetch group names and extract office locations from group names
                foreach (var gid in _config.AzureAd.AllowedGroupIds)
                {
                    try
                    {
                        // Get group info to extract office from group name
                        var group = await _graph.Groups[gid].GetAsync(r =>
                        {
                            r.QueryParameters.Select = new[] { "id", "displayName" };
                        });
                        
                        if (group?.DisplayName != null)
                        {
                            // Extract office from group name (e.g., "Maesteg - Staff" -> "Maesteg")
                            var groupName = group.DisplayName;
                            if (groupName.Contains(" - "))
                            {
                                var officeName = groupName.Split(" - ")[0].Trim();
                                if (!string.IsNullOrWhiteSpace(officeName) && !officeOptions.Contains(officeName, StringComparer.OrdinalIgnoreCase))
                                {
                                    officeOptions.Add(officeName);
                                }
                            }
                        }
                        
                        // Get group members
                        var gresp = await _graph.Groups[gid].Members.GraphUser.GetAsync(r =>
                        {
                            r.QueryParameters.Select = new[] { "id", "displayName", "givenName", "surname", "userPrincipalName", "jobTitle", "department", "officeLocation", "mail", "accountEnabled", "proxyAddresses" };
                            r.QueryParameters.Top = 999;
                        });
                        foreach (var u in (gresp?.Value ?? new List<User>()))
                        {
                            if (u?.Id != null && !union.ContainsKey(u.Id)) union[u.Id] = u;
                        }
                    }
                    catch { }
                }
                users = union.Values.ToList();
                
                // Sort office options
                officeOptions = officeOptions.OrderBy(s => s, StringComparer.OrdinalIgnoreCase).ToList();
            }
            else
            {
                // Pull a page of users and populate distinct department, jobTitle, office values (domain-restricted)
                var resp = await _graph.Users.GetAsync(r =>
                {
                    r.QueryParameters.Select = new[] { "department", "jobTitle", "officeLocation", "userPrincipalName", "mail", "accountEnabled", "displayName" };
                    r.QueryParameters.Filter = $"(endswith(userPrincipalName,'{AllowedDomain}') or endswith(mail,'{AllowedDomain}')) and (accountEnabled eq true) and (not contains(displayName,'admin') and not contains(userPrincipalName,'admin')) and (not contains(userPrincipalName,'.prod') and not contains(mail,'.prod')) and (not contains(userPrincipalName,'break.glass') and not contains(mail,'break.glass'))";
                    r.QueryParameters.Top = 200;
                    r.QueryParameters.Count = true;
                    r.Headers.Add("ConsistencyLevel", "eventual");
                });
                users = resp?.Value ?? new List<User>();
                
                // Use actual office locations from user profiles
                officeOptions = users.Select(u => u.OfficeLocation).Where(s => !string.IsNullOrWhiteSpace(s)).Cast<string>().Distinct(StringComparer.OrdinalIgnoreCase).OrderBy(s => s).ToList();
            }

            var depts = users.Select(u => u.Department).Where(s => !string.IsNullOrWhiteSpace(s)).Distinct(StringComparer.OrdinalIgnoreCase).OrderBy(s => s).ToList();
            var titles = users.Select(u => u.JobTitle).Where(s => !string.IsNullOrWhiteSpace(s)).Distinct(StringComparer.OrdinalIgnoreCase).OrderBy(s => s).ToList();

            if (OfficeBox is ComboBox officeCb) officeCb.ItemsSource = officeOptions;

            // Populate dept/title with ALL values when no office is selected (group mode)
            // and keep previous behavior for non-group mode
            if (DepartmentBox is ComboBox deptCb) deptCb.ItemsSource = depts;
            if (JobTitleBox is ComboBox titleCb) titleCb.ItemsSource = titles;
        }

        private async Task UpdateFilterOptionsAsync()
        {
            if (!UseGroupWhitelist) return; // Only for group whitelist mode

            // Cancel any in-flight option updates to avoid stale results overwriting newer selections
            _filterOptionsCts?.Cancel();
            _filterOptionsCts = new CancellationTokenSource();
            var token = _filterOptionsCts.Token;

            // Get current filter selections
            string? selectedOffice = null;
            string? selectedDepartment = null;
            string? selectedTitle = null;

            if (OfficeBox is ComboBox officeCb)
            {
                var val = (officeCb.SelectedItem as string) ?? officeCb.Text;
                if (!string.IsNullOrWhiteSpace(val)) selectedOffice = val.Trim();
            }
            if (DepartmentBox is ComboBox deptCb)
            {
                var val = (deptCb.SelectedItem as string) ?? deptCb.Text;
                if (!string.IsNullOrWhiteSpace(val)) selectedDepartment = val.Trim();
            }
            if (JobTitleBox is ComboBox titleCb)
            {
                var val = (titleCb.SelectedItem as string) ?? titleCb.Text;
                if (!string.IsNullOrWhiteSpace(val)) selectedTitle = val.Trim();
            }

            // Get filtered users based on current selections, but don't apply department/title filters when updating their options
            var filteredUsers = await GetFilteredUsersAsync(selectedOffice, null, null, token); // Only filter by office for updating dept/title options

            // Debug: Log what we found
            System.Diagnostics.Debug.WriteLine($"UpdateFilterOptionsAsync: Office='{selectedOffice}', Found {filteredUsers.Count} users");
            foreach (var user in filteredUsers.Take(5)) // Log first 5 users
            {
                System.Diagnostics.Debug.WriteLine($"  User: {user.DisplayName}, Dept: {user.Department}, Title: {user.JobTitle}");
            }

            // Update filter options based on filtered users
            var availableDepts = filteredUsers.Select(u => u.Department).Where(s => !string.IsNullOrWhiteSpace(s)).Cast<string>().Distinct(StringComparer.OrdinalIgnoreCase).OrderBy(s => s).ToList();
            var availableTitles = filteredUsers.Select(u => u.JobTitle).Where(s => !string.IsNullOrWhiteSpace(s)).Cast<string>().Distinct(StringComparer.OrdinalIgnoreCase).OrderBy(s => s).ToList();

            System.Diagnostics.Debug.WriteLine($"Available Departments: {string.Join(", ", availableDepts)}");
            System.Diagnostics.Debug.WriteLine($"Available Titles: {string.Join(", ", availableTitles)}");

            // If department is selected, filter titles by that department
            if (!string.IsNullOrEmpty(selectedDepartment))
            {
                var deptFilteredUsers = await GetFilteredUsersAsync(selectedOffice, selectedDepartment, null, token);
                availableTitles = deptFilteredUsers.Select(u => u.JobTitle).Where(s => !string.IsNullOrWhiteSpace(s)).Cast<string>().Distinct(StringComparer.OrdinalIgnoreCase).OrderBy(s => s).ToList();
            }

            // If title is selected, filter departments by that title
            if (!string.IsNullOrEmpty(selectedTitle))
            {
                var titleFilteredUsers = await GetFilteredUsersAsync(selectedOffice, null, selectedTitle, token);
                availableDepts = titleFilteredUsers.Select(u => u.Department).Where(s => !string.IsNullOrWhiteSpace(s)).Cast<string>().Distinct(StringComparer.OrdinalIgnoreCase).OrderBy(s => s).ToList();
            }

            // Update ComboBox sources while preserving current selections
            if (DepartmentBox is ComboBox deptBox)
            {
                var currentDept = deptBox.Text;
                deptBox.ItemsSource = availableDepts;
                if (!string.IsNullOrEmpty(currentDept) && availableDepts.Contains(currentDept, StringComparer.OrdinalIgnoreCase))
                    deptBox.Text = currentDept;
            }

            if (JobTitleBox is ComboBox jobBox)
            {
                var currentTitle = jobBox.Text;
                jobBox.ItemsSource = availableTitles;
                if (!string.IsNullOrEmpty(currentTitle) && availableTitles.Contains(currentTitle, StringComparer.OrdinalIgnoreCase))
                    jobBox.Text = currentTitle;
            }

        }

        private async Task<List<User>> GetFilteredUsersAsync(string? selectedOffice, string? selectedDepartment, string? selectedTitle, CancellationToken token)
        {
            var union = new Dictionary<string, User>(StringComparer.OrdinalIgnoreCase);

            System.Diagnostics.Debug.WriteLine($"GetFilteredUsersAsync: Office='{selectedOffice}', Dept='{selectedDepartment}', Title='{selectedTitle}'");

            // Cancel any in-flight filter updates to avoid stale results overwriting newer selections
            foreach (var gid in _config.AzureAd.AllowedGroupIds)
            {
                try
                {
                    // Get group info first
                    var group = await _graph.Groups[gid].GetAsync(r =>
                    {
                        r.QueryParameters.Select = new[] { "id", "displayName" };
                    }, token);

                    System.Diagnostics.Debug.WriteLine($"Processing group: {group?.DisplayName}");

                    // If office filter is selected, only include groups that match the office
                    if (!string.IsNullOrEmpty(selectedOffice))
                    {
                        if (group?.DisplayName != null)
                        {
                            var groupName = group.DisplayName;
                            if (groupName.Contains(" - "))
                            {
                                var officeName = groupName.Split(" - ")[0].Trim();
                                if (!officeName.Equals(selectedOffice, StringComparison.OrdinalIgnoreCase))
                                {
                                    System.Diagnostics.Debug.WriteLine($"  Skipping group '{groupName}' - office '{officeName}' doesn't match '{selectedOffice}'");
                                    continue; // Skip this group as it doesn't match the selected office
                                }
                                else
                                {
                                    System.Diagnostics.Debug.WriteLine($"  Including group '{groupName}' - office '{officeName}' matches '{selectedOffice}'");
                                }
                            }
                            else
                            {
                                System.Diagnostics.Debug.WriteLine($"  Skipping group '{groupName}' - doesn't follow 'Office - Type' pattern");
                                continue; // Skip groups that don't follow the "Office - Type" pattern
                            }
                        }
                    }

                    var gresp = await _graph.Groups[gid].Members.GraphUser.GetAsync(r =>
                    {
                        r.QueryParameters.Select = new[] { "id", "displayName", "userPrincipalName", "jobTitle", "department", "officeLocation", "mail", "accountEnabled", "proxyAddresses" };
                        r.QueryParameters.Top = 999;
                    }, token);

                    var groupUsers = gresp?.Value ?? new List<User>();
                    System.Diagnostics.Debug.WriteLine($"  Found {groupUsers.Count} users in group");

                    foreach (var u in groupUsers)
                    {
                        if (u?.Id != null && !union.ContainsKey(u.Id)) 
                        {
                            union[u.Id] = u;
                            System.Diagnostics.Debug.WriteLine($"    Added user: {u.DisplayName}, Dept: {u.Department}, Title: {u.JobTitle}");
                        }
                    }
                }
                catch { }
            }

            // Apply client-side filtering
            IEnumerable<User> queryable = union.Values
                .Where(u => (u.AccountEnabled ?? true))
                .Where(u => !(u.DisplayName?.Contains("admin", StringComparison.OrdinalIgnoreCase) ?? false))
                .Where(u => !(u.UserPrincipalName?.Contains("admin", StringComparison.OrdinalIgnoreCase) ?? false))
                .Where(u => !(u.UserPrincipalName?.Contains(".prod", StringComparison.OrdinalIgnoreCase) ?? false))
                .Where(u => !(u.Mail?.Contains(".prod", StringComparison.OrdinalIgnoreCase) ?? false))
                .Where(u => !(u.UserPrincipalName?.Contains("break.glass", StringComparison.OrdinalIgnoreCase) ?? false))
                .Where(u => !(u.Mail?.Contains("break.glass", StringComparison.OrdinalIgnoreCase) ?? false))
                .Where(u => (u.UserPrincipalName?.EndsWith(AllowedDomain, StringComparison.OrdinalIgnoreCase) ?? false) || (u.Mail?.EndsWith(AllowedDomain, StringComparison.OrdinalIgnoreCase) ?? false));

            // Apply department filter if selected
            if (!string.IsNullOrEmpty(selectedDepartment))
                queryable = queryable.Where(u => u.Department?.StartsWith(selectedDepartment, StringComparison.OrdinalIgnoreCase) ?? false);

            // Apply title filter if selected
            if (!string.IsNullOrEmpty(selectedTitle))
                queryable = queryable.Where(u => u.JobTitle?.StartsWith(selectedTitle, StringComparison.OrdinalIgnoreCase) ?? false);

            return queryable.ToList();
        }

        private async Task<SKBitmap?> BuildBadgeAsync(User user)
        {
            SKBitmap? photoBmp = null;
            try
            {
                var stream = await _graph.Users[user.Id].Photo.Content.GetAsync();
                if (stream != null)
                {
                    using var ms = new MemoryStream();
                    await stream.CopyToAsync(ms);
                    ms.Position = 0;
                    photoBmp = BadgeRenderer.LoadSkBitmapFromStream(ms);
                }
            }
            catch { }

            var displayName = BuildDisplayName(user);
            var upn = user.UserPrincipalName ?? string.Empty;
            var role = user.JobTitle ?? string.Empty;
            var dept = user.Department ?? string.Empty;
            var proxyAddresses = user.ProxyAddresses ?? new List<string>();
            return _renderer.Render(displayName, role, dept, upn, proxyAddresses, photoBmp);
        }

        private async void OnPrintSelectedClick(object sender, RoutedEventArgs e)
        {
            try
            {
                var selected = ResultsList.SelectedItems?.OfType<UserItem>().ToList() ?? new List<UserItem>();
                if (selected.Count == 0)
                {
                    StatusText.Text = "No users selected";
                    return;
                }

                // Confirm batch print
                var confirm = MessageBox.Show(
                    $"You are about to print {selected.Count} Siderise security badge(s).\n\nDo you want to continue?",
                    "Siderise Security Badge Printer - Confirm Batch Print",
                    MessageBoxButton.YesNo,
                    MessageBoxImage.Question);
                if (confirm != MessageBoxResult.Yes)
                {
                    StatusText.Text = "Batch print cancelled";
                    return;
                }

                StatusText.Text = $"Batch printing {selected.Count} badge(s)...";
                foreach (var it in selected)
                {
                    var badge = await BuildBadgeAsync(it.User);
                    if (badge != null) _printer.Print(badge);
                }
                StatusText.Text = $"Printed {selected.Count} badge(s)";
            }
            catch (Exception ex)
            {
                StatusText.Text = $"Print error: {ex.Message}";
            }
        }

        private async void OnSearchClick(object sender, RoutedEventArgs e)
        {
            var q = (SearchBox.Text ?? string.Empty).Trim();
            try
            {
                StatusText.Text = "Searching...";
                await QueryAndBindUsersAsync(q);
                StatusText.Text = "Search complete";
            }
            catch (Exception ex)
            {
                StatusText.Text = $"Error: {ex.Message}";
            }
        }

        private async void OnFilterTextChanged(object sender, TextChangedEventArgs e)
        {
            // Only update filter options on ComboBox (Office) text edits; SearchBox should not refresh option lists
            if (sender is ComboBox)
            {
                await UpdateFilterOptionsAsync();
            }
            await DebouncedFilterAsync();
        }

        private async void OnFilterSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            // Update filter options based on which control changed to avoid fighting user selection
            if (UseGroupWhitelist)
            {
                if (ReferenceEquals(sender, OfficeBox))
                {
                    await UpdateFilterOptionsAsync();
                }
                else if (ReferenceEquals(sender, DepartmentBox))
                {
                    // If dropdown still open, defer updates until closed
                    if (DepartmentBox is ComboBox cb && cb.IsDropDownOpen) return;
                    await UpdateTitlesForCurrentFiltersAsync();
                }
                else if (ReferenceEquals(sender, JobTitleBox))
                {
                    // If dropdown still open, defer updates until closed
                    if (JobTitleBox is ComboBox cb && cb.IsDropDownOpen) return;
                    await UpdateDepartmentsForCurrentFiltersAsync();
                }
            }

            await QueryAndBindUsersAsync(SearchBox.Text?.Trim() ?? string.Empty);
        }

        private async void OnDepartmentClosed(object sender, EventArgs e)
        {
            if (DepartmentBox is ComboBox dept && dept.SelectedItem is string sel && !string.IsNullOrWhiteSpace(sel))
            {
                // Commit title options based on chosen department
                await UpdateTitlesForCurrentFiltersAsync();
                await QueryAndBindUsersAsync(SearchBox.Text?.Trim() ?? string.Empty);
            }
        }

        private async void OnTitleClosed(object sender, EventArgs e)
        {
            if (JobTitleBox is ComboBox title && title.SelectedItem is string sel && !string.IsNullOrWhiteSpace(sel))
            {
                // Commit department options based on chosen title
                await UpdateDepartmentsForCurrentFiltersAsync();
                await QueryAndBindUsersAsync(SearchBox.Text?.Trim() ?? string.Empty);
            }
        }

        private async Task UpdateTitlesForCurrentFiltersAsync()
        {
            if (!UseGroupWhitelist) return;
            _filterOptionsCts?.Cancel();
            _filterOptionsCts = new CancellationTokenSource();
            var token = _filterOptionsCts.Token;

            string? selectedOffice = null;
            string? selectedDepartment = null;

            if (OfficeBox is ComboBox officeCb)
            {
                var val = (officeCb.SelectedItem as string) ?? officeCb.Text;
                if (!string.IsNullOrWhiteSpace(val)) selectedOffice = val.Trim();
            }
            if (DepartmentBox is ComboBox deptCb)
            {
                var val = (deptCb.SelectedItem as string) ?? deptCb.Text;
                if (!string.IsNullOrWhiteSpace(val)) selectedDepartment = val.Trim();
            }

            var users = await GetFilteredUsersAsync(selectedOffice, selectedDepartment, null, token);
            var availableTitles = users.Select(u => u.JobTitle).Where(s => !string.IsNullOrWhiteSpace(s)).Cast<string>().Distinct(StringComparer.OrdinalIgnoreCase).OrderBy(s => s).ToList();

            if (JobTitleBox is ComboBox jobBox)
            {
                var currentTitle = jobBox.Text;
                jobBox.ItemsSource = availableTitles;
                if (!string.IsNullOrEmpty(currentTitle) && availableTitles.Contains(currentTitle, StringComparer.OrdinalIgnoreCase))
                    jobBox.SelectedItem = availableTitles.First(t => string.Equals(t, currentTitle, StringComparison.OrdinalIgnoreCase));
            }
        }

        private async Task UpdateDepartmentsForCurrentFiltersAsync()
        {
            if (!UseGroupWhitelist) return;
            _filterOptionsCts?.Cancel();
            _filterOptionsCts = new CancellationTokenSource();
            var token = _filterOptionsCts.Token;

            string? selectedOffice = null;
            string? selectedTitle = null;

            if (OfficeBox is ComboBox officeCb)
            {
                var val = (officeCb.SelectedItem as string) ?? officeCb.Text;
                if (!string.IsNullOrWhiteSpace(val)) selectedOffice = val.Trim();
            }
            if (JobTitleBox is ComboBox titleCb)
            {
                var val = (titleCb.SelectedItem as string) ?? titleCb.Text;
                if (!string.IsNullOrWhiteSpace(val)) selectedTitle = val.Trim();
            }

            var users = await GetFilteredUsersAsync(selectedOffice, null, selectedTitle, token);
            var availableDepts = users.Select(u => u.Department).Where(s => !string.IsNullOrWhiteSpace(s)).Cast<string>().Distinct(StringComparer.OrdinalIgnoreCase).OrderBy(s => s).ToList();

            if (DepartmentBox is ComboBox deptBox)
            {
                var currentDept = deptBox.Text;
                deptBox.ItemsSource = availableDepts;
                if (!string.IsNullOrEmpty(currentDept) && availableDepts.Contains(currentDept, StringComparer.OrdinalIgnoreCase))
                    deptBox.SelectedItem = availableDepts.First(d => string.Equals(d, currentDept, StringComparison.OrdinalIgnoreCase));
            }
        }

        private async Task DebouncedFilterAsync(int delayMs = 300)
        {
            _filterCts?.Cancel();
            _filterCts = new CancellationTokenSource();
            var token = _filterCts.Token;
            try
            {
                await Task.Delay(delayMs, token);
                if (!token.IsCancellationRequested)
                {
                    await QueryAndBindUsersAsync(SearchBox.Text?.Trim() ?? string.Empty);
                }
            }
            catch (TaskCanceledException) { }
        }

        private async Task QueryAndBindUsersAsync(string query)
        {
            var safe = (query ?? string.Empty).Replace("'", "''");

            // Base filters (non-group mode)
            var filters = new List<string>
            {
                $"(endswith(userPrincipalName,'{AllowedDomain}') or endswith(mail,'{AllowedDomain}'))",
                "(accountEnabled eq true)",
                $"(not contains(displayName,'admin') and not contains(userPrincipalName,'admin'))",
                $"(not contains(userPrincipalName,'.prod') and not contains(mail,'.prod'))",
                $"(not contains(userPrincipalName,'break.glass') and not contains(mail,'break.glass'))"
            };
            if (!string.IsNullOrWhiteSpace(safe))
            {
                filters.Add($"(startswith(displayName,'{safe}') or startswith(userPrincipalName,'{safe}'))");
            }
            if (DepartmentBox is ComboBox deptCb)
            {
                var deptVal = (deptCb.SelectedItem as string) ?? deptCb.Text;
                if (!string.IsNullOrWhiteSpace(deptVal))
                {
                    var d = deptVal.Replace("'", "''");
                    filters.Add($"(startswith(department,'{d}'))");
                }
            }
            if (JobTitleBox is ComboBox titleCb)
            {
                var titleVal = (titleCb.SelectedItem as string) ?? titleCb.Text;
                if (!string.IsNullOrWhiteSpace(titleVal))
                {
                    var t = titleVal.Replace("'", "''");
                    filters.Add($"(startswith(jobTitle,'{t}'))");
                }
            }
            // Office filtering is handled by group filtering in group whitelist mode
            if (!UseGroupWhitelist && OfficeBox is ComboBox officeCb)
            {
                var officeVal = (officeCb.SelectedItem as string) ?? officeCb.Text;
                if (!string.IsNullOrWhiteSpace(officeVal))
                {
                    var o = officeVal.Replace("'", "''");
                    filters.Add($"(startswith(officeLocation,'{o}'))");
                }
            }

            List<User> users;
            if (UseGroupWhitelist)
            {
                // Check if office filter is selected to filter by specific groups
                string? selectedOffice = null;
                if (OfficeBox is ComboBox officeCombo)
                {
                    var val = (officeCombo.SelectedItem as string) ?? officeCombo.Text;
                    if (!string.IsNullOrWhiteSpace(val)) selectedOffice = val.Trim();
                }

                var union = new Dictionary<string, User>(StringComparer.OrdinalIgnoreCase);
                foreach (var gid in _config.AzureAd.AllowedGroupIds)
                {
                    try
                    {
                        // If office filter is selected, only include groups that match the office
                        if (!string.IsNullOrEmpty(selectedOffice))
                        {
                            var group = await _graph.Groups[gid].GetAsync(r =>
                            {
                                r.QueryParameters.Select = new[] { "id", "displayName" };
                            });
                            
                            if (group?.DisplayName != null)
                            {
                                var groupName = group.DisplayName;
                                if (groupName.Contains(" - "))
                                {
                                    var officeName = groupName.Split(" - ")[0].Trim();
                                    if (!officeName.Equals(selectedOffice, StringComparison.OrdinalIgnoreCase))
                                    {
                                        continue; // Skip this group as it doesn't match the selected office
                                    }
                                }
                                else
                                {
                                    continue; // Skip groups that don't follow the "Office - Type" pattern
                                }
                            }
                        }

                        var gresp = await _graph.Groups[gid].Members.GraphUser.GetAsync(r =>
                        {
                            r.QueryParameters.Select = new[] { "id", "displayName", "userPrincipalName", "jobTitle", "department", "officeLocation", "mail", "accountEnabled", "proxyAddresses" };
                            r.QueryParameters.Top = 999;
                        });
                        foreach (var u in (gresp?.Value ?? new List<User>()))
                        {
                            if (u?.Id != null && !union.ContainsKey(u.Id)) union[u.Id] = u;
                        }
                    }
                    catch { }
                }
                users = union.Values.ToList();
            }
            else
            {
                var filter = string.Join(" and ", filters);
                var resp = await _graph.Users.GetAsync(r =>
                {
                    r.QueryParameters.Select = new[] { "id", "displayName", "userPrincipalName", "jobTitle", "department", "officeLocation", "mail", "accountEnabled", "proxyAddresses" };
                    r.QueryParameters.Filter = filter;
                    r.QueryParameters.Orderby = new[] { "displayName" };
                    r.QueryParameters.Top = 50;
                    r.QueryParameters.Count = true;
                    r.Headers.Add("ConsistencyLevel", "eventual");
                });
                users = resp?.Value?.ToList() ?? new List<User>();
            }

            // Client-side enforcement (both modes)
            IEnumerable<User> queryable = users
                .Where(u => (u.AccountEnabled ?? true))
                .Where(u => !(u.DisplayName?.Contains("admin", StringComparison.OrdinalIgnoreCase) ?? false))
                .Where(u => !(u.UserPrincipalName?.Contains("admin", StringComparison.OrdinalIgnoreCase) ?? false))
                .Where(u => !(u.UserPrincipalName?.Contains(".prod", StringComparison.OrdinalIgnoreCase) ?? false))
                .Where(u => !(u.Mail?.Contains(".prod", StringComparison.OrdinalIgnoreCase) ?? false))
                .Where(u => !(u.UserPrincipalName?.Contains("break.glass", StringComparison.OrdinalIgnoreCase) ?? false))
                .Where(u => !(u.Mail?.Contains("break.glass", StringComparison.OrdinalIgnoreCase) ?? false))
                .Where(u => (u.UserPrincipalName?.EndsWith(AllowedDomain, StringComparison.OrdinalIgnoreCase) ?? false) || (u.Mail?.EndsWith(AllowedDomain, StringComparison.OrdinalIgnoreCase) ?? false));

            if (!string.IsNullOrWhiteSpace(safe))
                queryable = queryable.Where(u => (u.DisplayName?.StartsWith(safe, StringComparison.OrdinalIgnoreCase) ?? false) || (u.UserPrincipalName?.StartsWith(safe, StringComparison.OrdinalIgnoreCase) ?? false));
            if (DepartmentBox is ComboBox dcb && !string.IsNullOrWhiteSpace(dcb.Text))
                queryable = queryable.Where(u => u.Department?.StartsWith(dcb.Text, StringComparison.OrdinalIgnoreCase) ?? false);
            if (JobTitleBox is ComboBox tcb && !string.IsNullOrWhiteSpace(tcb.Text))
                queryable = queryable.Where(u => u.JobTitle?.StartsWith(tcb.Text, StringComparison.OrdinalIgnoreCase) ?? false);
            // Office filtering is handled by group membership in group whitelist mode
            if (!UseGroupWhitelist && OfficeBox is ComboBox ocb && !string.IsNullOrWhiteSpace(ocb.Text))
                queryable = queryable.Where(u => u.OfficeLocation?.StartsWith(ocb.Text, StringComparison.OrdinalIgnoreCase) ?? false);

            var items = queryable
                .OrderBy(u => u.DisplayName)
                .Select(u => new UserItem(u))
                .ToList();

            ResultsList.ItemsSource = items;
            // Reset preview state
            if (SinglePreviewBox != null) SinglePreviewBox.Visibility = Visibility.Collapsed;
            PreviewImage.Source = null;
            PreviewImage.Visibility = Visibility.Collapsed;
            if (PreviewItems != null) { PreviewItems.ItemsSource = null; PreviewItems.Visibility = Visibility.Collapsed; }
            if (PreviewNamesPanel != null) { PreviewNamesPanel.ItemsSource = null; }
            if (PreviewNamesHeader != null) { PreviewNamesHeader.Text = string.Empty; }
            if (PreviewNamesContainer != null) { PreviewNamesContainer.Visibility = Visibility.Collapsed; }
            if (PreviewFrame != null) PreviewFrame.BorderThickness = new Thickness(0);
            PrintButton.IsEnabled = false;
            PrintSelectedButton.IsEnabled = false;
        }

        private void OnResultsListPreviewMouseWheel(object sender, MouseWheelEventArgs e)
        {
            // Forward mouse wheel events to the parent ScrollViewer for smooth scrolling
            if (!e.Handled && ResultsScrollViewer != null)
            {
                e.Handled = true;
                var eventArg = new MouseWheelEventArgs(e.MouseDevice, e.Timestamp, e.Delta)
                {
                    RoutedEvent = UIElement.MouseWheelEvent,
                    Source = sender
                };
                ResultsScrollViewer.RaiseEvent(eventArg);
            }
        }

        private async void OnResultSelected(object sender, SelectionChangedEventArgs e)
        {
            // Enable Batch Print only when more than one item is selected
            var selCount = ResultsList.SelectedItems?.Count ?? 0;
            PrintSelectedButton.IsEnabled = selCount > 1;

            // Reset preview visuals
            if (SinglePreviewBox != null) SinglePreviewBox.Visibility = Visibility.Collapsed;
            PreviewImage.Source = null;
            PreviewImage.Visibility = Visibility.Collapsed;
            if (PreviewItems != null) { PreviewItems.ItemsSource = null; PreviewItems.Visibility = Visibility.Collapsed; }
            if (PreviewNamesPanel != null) { PreviewNamesPanel.ItemsSource = null; }
            if (PreviewNamesHeader != null) { PreviewNamesHeader.Text = string.Empty; }
            if (PreviewNamesContainer != null) { PreviewNamesContainer.Visibility = Visibility.Collapsed; }
            if (PreviewFrame != null) PreviewFrame.BorderThickness = new Thickness(0);

            if (selCount == 0)
            {
                PrintButton.IsEnabled = false;
                return;
            }

            // Single selection: show single preview image centered & scaled
            if (selCount == 1)
            {
                if (ResultsList.SelectedItem is not UserItem item)
                {
                    PrintButton.IsEnabled = false;
                    return;
                }
                _selectedUser = item.User;
                try
                {
                    StatusText.Text = "Loading photo and rendering preview...";
                    var badge = await BuildBadgeAsync(_selectedUser);
                    if (badge != null)
                    {
                        _lastBadge = badge;
                        PreviewImage.Source = ToPreviewBitmapImage(_lastBadge);
                        PreviewImage.Visibility = Visibility.Visible;
                        if (SinglePreviewBox != null) SinglePreviewBox.Visibility = Visibility.Visible;
                        if (PreviewFrame != null)
                        {
                            PreviewFrame.BorderThickness = new Thickness(2); // show frame only for single preview
                            PreviewFrame.VerticalAlignment = VerticalAlignment.Center;
                            PreviewFrame.Margin = new Thickness(0);
                        }
                        ConstrainSinglePreviewToViewport();
                        PrintButton.IsEnabled = true;
                        StatusText.Text = "Preview ready";
                    }
                    else
                    {
                        StatusText.Text = "Preview not available";
                    }
                }
                catch (Exception ex)
                {
                    StatusText.Text = $"Error: {ex.Message}";
                }
                return;
            }

            // Multiple selection: up to 6 previews, else list names
            try
            {
                var selectedItems = (ResultsList.SelectedItems?.OfType<UserItem>() ?? System.Linq.Enumerable.Empty<UserItem>()).ToList();
                PrintButton.IsEnabled = false; // disable single print in multi mode

                if (selectedItems.Count <= 6)
                {
                    StatusText.Text = $"Rendering {selectedItems.Count} previews...";
                    var bitmaps = new List<BitmapImage>();
                    foreach (var it in selectedItems)
                    {
                        var bmp = await BuildBadgeAsync(it.User);
                        if (bmp != null)
                        {
                            // Use the full image in multi-preview to avoid any edge clipping
                            bitmaps.Add(ToBitmapImage(bmp));
                        }
                    }
                    if (PreviewItems != null)
                    {
                        PreviewItems.ItemsSource = bitmaps;
                        PreviewItems.Visibility = Visibility.Visible;
                        // Also clamp height immediately based on current viewport to prevent bottom clipping
                        if (PreviewScroller != null)
                        {
                            var vh = PreviewScroller.ViewportHeight;
                            if (vh > 0) PreviewItems.MaxHeight = Math.Max(0, vh - 20);
                        }
                    }
                    if (PreviewFrame != null)
                    {
                        PreviewFrame.BorderThickness = new Thickness(0); // no outer frame in multi
                        PreviewFrame.VerticalAlignment = VerticalAlignment.Top; // move multi preview up
                        PreviewFrame.Margin = new Thickness(0);
                                            }
                    // In multi mode, force content to fit viewport width (no horizontal scroll)
                    if (PreviewScroller != null)
                    {
                        PreviewScroller.HorizontalScrollBarVisibility = System.Windows.Controls.ScrollBarVisibility.Disabled;
                        PreviewScroller.VerticalScrollBarVisibility = System.Windows.Controls.ScrollBarVisibility.Auto;
                    }
                                        StatusText.Text = "Previews ready";
                }
                else
                {
                    var names = selectedItems.Select(si => BuildDisplayName(si.User)).ToList();
                    if (PreviewNamesContainer != null && PreviewNamesHeader != null && PreviewNamesPanel != null)
                    {
                        PreviewNamesHeader.Text = $"Printing {names.Count} badges:";
                        PreviewNamesPanel.ItemsSource = names.Select(n => $"- {n}");
                        PreviewNamesContainer.Visibility = Visibility.Visible;
                        if (PreviewFrame != null)
                        {
                            PreviewFrame.BorderThickness = new Thickness(0);
                            PreviewFrame.VerticalAlignment = VerticalAlignment.Top;
                            PreviewFrame.Margin = new Thickness(0);
                        }
                        // Size the names area to the viewport so items flow into vertical columns from the left
                        if (PreviewScroller != null && PreviewNamesScroll != null)
                        {
                            var vw = PreviewScroller.ViewportWidth;
                            var vh = PreviewScroller.ViewportHeight;
                            if (vw > 0) PreviewNamesScroll.Width = vw - 24;
                        }
                        // Horizontal scroll to see more columns; no vertical scroll in list mode
                        if (PreviewScroller != null)
                        {
                            PreviewScroller.HorizontalScrollBarVisibility = System.Windows.Controls.ScrollBarVisibility.Auto;
                            PreviewScroller.VerticalScrollBarVisibility = System.Windows.Controls.ScrollBarVisibility.Disabled;
                        }
                    }
                    if (PreviewFrame != null) PreviewFrame.BorderThickness = new Thickness(0); // no frame in list view
                    StatusText.Text = "Too many to preview";
                }
            }
            catch (Exception ex)
            {
                StatusText.Text = $"Error: {ex.Message}";
            }
        }

        private void OnPrintClick(object sender, RoutedEventArgs e)
        {
            if (_lastBadge == null)
            {
                StatusText.Text = "No badge to print";
                return;
            }

            try
            {
                _printer.Print(_lastBadge);
                StatusText.Text = $"Sent to printer '{_config.Printer.Name}'";
            }
            catch (Exception ex)
            {
                StatusText.Text = $"Print error: {ex.Message}";
            }
        }

        private async void OnClearFiltersClick(object sender, RoutedEventArgs e)
        {
            SearchBox.Text = string.Empty;
            if (DepartmentBox is ComboBox d) { d.SelectedItem = null; d.Text = string.Empty; }
            if (JobTitleBox is ComboBox t) { t.SelectedItem = null; t.Text = string.Empty; }
            if (OfficeBox is ComboBox o) { o.SelectedItem = null; o.Text = string.Empty; }
            PreviewImage.Source = null;
            PrintButton.IsEnabled = false;
            PrintSelectedButton.IsEnabled = false;
            
            // Reload all filter options
            await LoadFilterValuesAsync();
            await QueryAndBindUsersAsync(string.Empty);
        }

        private static string BuildDisplayName(User user)
        {
            // First + Last only
            var given = user.GivenName;
            var sur = user.Surname;
            if (!string.IsNullOrWhiteSpace(given) && !string.IsNullOrWhiteSpace(sur))
                return $"{given} {sur}";
            // Do not fall back to UPN/email
            return user.DisplayName ?? "Unknown";
        }

        private static BitmapImage ToBitmapImage(SKBitmap bmp)
        {
            using var data = bmp.Encode(SKEncodedImageFormat.Png, 100);
            using var ms = new MemoryStream(data.ToArray());
            var image = new BitmapImage();
            image.BeginInit();
            image.CacheOption = BitmapCacheOption.OnLoad;
            image.StreamSource = ms;
            image.EndInit();
            image.Freeze();
            return image;
        }

        private static BitmapImage ToPreviewBitmapImage(SKBitmap bmp)
        {
            var cropped = CropTransparent(bmp);
            try
            {
                var src = cropped ?? bmp;
                return ToBitmapImage(src);
            }
            finally
            {
                cropped?.Dispose();
            }
        }

        private static SKBitmap? CropTransparent(SKBitmap bmp)
        {
            // Scan for non-transparent bounds
            int minX = bmp.Width, minY = bmp.Height, maxX = -1, maxY = -1;
            for (int y = 0; y < bmp.Height; y++)
            {
                for (int x = 0; x < bmp.Width; x++)
                {
                    var c = bmp.GetPixel(x, y);
                    if (c.Alpha != 0)
                    {
                        if (x < minX) minX = x;
                        if (y < minY) minY = y;
                        if (x > maxX) maxX = x;
                        if (y > maxY) maxY = y;
                    }
                }
            }
            if (maxX < minX || maxY < minY) return null; // all transparent

            int w = maxX - minX + 1;
            int h = maxY - minY + 1;
            var cropped = new SKBitmap(w, h, bmp.ColorType, bmp.AlphaType);
            using (var canvas = new SKCanvas(cropped))
            {
                var src = new SKRectI(minX, minY, maxX + 1, maxY + 1);
                var dst = new SKRectI(0, 0, w, h);
                canvas.DrawBitmap(bmp, src, dst);
                canvas.Flush();
            }
            return cropped;
        }

        private class UserItem
        {
            public User User { get; }
            public string Display { get; }
            public UserItem(User user)
            {
                User = user;
                var name = BuildDisplayName(user);
                // Do not show email/UPN in lists
                Display = name;
            }
        }
    }
}
