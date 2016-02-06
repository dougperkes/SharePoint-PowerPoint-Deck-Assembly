<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Combine.aspx.cs" Inherits="PowerPointAssemblyWeb.Combine" %>
<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>Combine Decks</title>
    <script type="text/javascript" src="Scripts/jquery-2.2.0.min.js"></script>
    <script type="text/javascript" src="Scripts/UXScripts.js"></script>
    <script type="text/javascript">
        $(document).ready(function () {
            $("#btnOk").click(function (event) {
                $("#divWaitingPanel").show();
            });

            $(".cboOrder").change(function (args) {
                //get the new value and adjust everything above it
                var target = $(args.target);
                var told = parseInt(target.attr("data-prev"));
                var tnew = parseInt(target.val());
                var tindex = $(".cboOrder").index(target);
                target.attr("data-prev", target.val());
                $(".cboOrder").each(function (i, e) {
                    var item = $(e);

                    //ignore the one that triggered the event
                    if (i != tindex) {
                        var val = parseInt($(e).val());
                        if ((tnew < told) && (val >= tnew)) {
                            //add to item value
                            item.val(val + 1 + '');
                            item.attr("data-prev", item.val());
                        }
                        else if ((tnew > told) && (val <= tnew)) {
                            //subtract from item value
                            item.val(val - 1 + '');
                            item.attr("data-prev", item.val());
                        }
                    }
                });
            });
        });
    </script>
</head>
<body style="display: none;">
    <form id="form1" runat="server">
    <div style="width: 100%;height:400px; overflow:auto">
        <div id="divWaitingPanel" style="position: absolute; z-index: 3; background: rgb(255, 255, 255); left: 0px; right: 0px; top: 0px; bottom: 0px; display: none;">
            <div style="text-align: center;">
                <img alt="Working on it" src="data:image/gif;base64,R0lGODlhEAAQAIAAAFLOQv///yH/C05FVFNDQVBFMi4wAwEAAAAh+QQFCgABACwJAAIAAgACAAACAoRRACH5BAUKAAEALAwABQACAAIAAAIChFEAIfkEBQoAAQAsDAAJAAIAAgAAAgKEUQAh+QQFCgABACwJAAwAAgACAAACAoRRACH5BAUKAAEALAUADAACAAIAAAIChFEAIfkEBQoAAQAsAgAJAAIAAgAAAgKEUQAh+QQFCgABACwCAAUAAgACAAACAoRRACH5BAkKAAEALAIAAgAMAAwAAAINjAFne8kPo5y02ouzLQAh+QQJCgABACwCAAIADAAMAAACF4wBphvID1uCyNEZM7Ov4v1p0hGOZlAAACH5BAkKAAEALAIAAgAMAAwAAAIUjAGmG8gPW4qS2rscRPp1rH3H1BUAIfkECQoAAQAsAgACAAkADAAAAhGMAaaX64peiLJa6rCVFHdQAAAh+QQJCgABACwCAAIABQAMAAACDYwBFqiX3mJjUM63QAEAIfkECQoAAQAsAgACAAUACQAAAgqMARaol95iY9AUACH5BAkKAAEALAIAAgAFAAUAAAIHjAEWqJeuCgAh+QQJCgABACwFAAIAAgACAAACAoRRADs=" style="width: 32px; height: 32px;" />
                <span class="ms-accentText" style="font-size: 36px;">&nbsp;Working on it...</span>
            </div>
        </div>
        <div style="width: 100%;">
            <h3 class="ms-core-form-line">Name the combined presentation</h3>
            <div class="ms-core-form-line">
                <asp:TextBox ID="txtFileName" runat="server" CssClass="ms-fullWidth"></asp:TextBox>
            </div>
        </div>
        <div style="width: 100%;">
            <asp:GridView ID="gridViewSelectedFiles" runat="server" AutoGenerateColumns="false" CssClass="ms-listviewtable" Width="100%" GridLines="None" OnRowDataBound="gridViewSelectedFiles_RowDataBound">
                <Columns>
                    <asp:TemplateField HeaderText="Include" ItemStyle-CssClass="ms-cellstyle ms-vb2" HeaderStyle-CssClass="ms-vh" ItemStyle-Width="40px" ItemStyle-HorizontalAlign="Center">
                        <ItemTemplate>
                            <asp:HiddenField ID="hdnItemId" runat="server" Value='<%# DataBinder.Eval(Container.DataItem, "Id") %>' />
                            <asp:Checkbox ID="chkItem" runat="server" Checked="true" />
                        </ItemTemplate>
                    </asp:TemplateField>
                    <asp:BoundField DataField="Name" HeaderText="Name" ItemStyle-CssClass="ms-cellstyle ms-vb2" HeaderStyle-CssClass="ms-vh" />
                    <asp:TemplateField HeaderText="Order" ItemStyle-CssClass="ms-cellstyle ms-vb2" HeaderStyle-CssClass="ms-vh" ItemStyle-Width="40px" ItemStyle-HorizontalAlign="Right">
                        <ItemTemplate>
                            <asp:DropDownList ID="cboOrder" runat="server" CssClass="cboOrder"></asp:DropDownList>
                        </ItemTemplate>
                    </asp:TemplateField>
                    <asp:TemplateField HeaderText="Keep Format" ItemStyle-CssClass="ms-cellstyle ms-vb2" HeaderStyle-CssClass="ms-vh" ItemStyle-Width="40px" ItemStyle-HorizontalAlign="Center">
                        <ItemTemplate>
                            <asp:Checkbox ID="chkFormat" runat="server" Checked="true" />
                        </ItemTemplate>
                    </asp:TemplateField>
                </Columns>
                <HeaderStyle CssClass="ms-viewheadertr ms-vhltr" />
                <RowStyle CssClass="ms-itmHoverEnabled ms-itmhover" />
                <AlternatingRowStyle CssClass="ms-alternating  ms-itmHoverEnabled ms-itmhover" />
                <EmptyDataTemplate>
                    <div style="color: red;">No presentations were selected</div>
                </EmptyDataTemplate>
            </asp:GridView>
            <asp:HiddenField ID="allIdsHidden" runat="server" />
        </div>
        <div style="float: right; width: 100%; text-align: right;">
            <asp:Button ID="btnOk" runat="server" Text="Ok" OnClick="btnOk_Click" />
            <button onclick="closeParentDialog(false); return false;">Cancel</button>
        </div>
    </div>
    </form>
    
</body>
</html>
